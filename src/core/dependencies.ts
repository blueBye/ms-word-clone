import {
  D2,
  MultiSet,
  MessageType,
  innerJoin,
  output,
} from "@electric-sql/d2ts";

/** Incremental style-to-paragraph dependencies. Both input frontiers advance on every commit. */
export class StyleDependencies {
  private graph = new D2({ initialFrontier: 0 });
  private edges = this.graph.newInput<[string, string]>();
  private styles = this.graph.newInput<[string, number]>();
  private membership = new Map<string, Set<string>>();
  private versions = new Map<string, number>();
  private edgeChanges: [[string, string], number][] = [];
  private styleChanges: [[string, number], number][] = [];
  private affected = new Set<string>();
  private version = 0;
  constructor() {
    this.edges.pipe(
      innerJoin(this.styles),
      output((message) => {
        if (message.type === MessageType.DATA)
          for (const [[, [id]], weight] of message.data.collection.getInner())
            if (weight) this.affected.add(id);
      }),
    );
    this.graph.finalize();
  }
  set(id: string, dependencies: Iterable<string>) {
    const next = new Set(dependencies),
      previous = this.membership.get(id) ?? new Set<string>();
    for (const style of previous)
      if (!next.has(style)) this.edgeChanges.push([[style, id], -1]);
    for (const style of next) {
      if (!previous.has(style)) this.edgeChanges.push([[style, id], 1]);
      if (!this.versions.has(style)) {
        this.versions.set(style, 0);
        this.styleChanges.push([[style, 0], 1]);
      }
    }
    if (next.size) this.membership.set(id, next);
    else this.membership.delete(id);
  }
  touch(ids: Iterable<string>) {
    for (const id of new Set(ids)) {
      const previous = this.versions.get(id);
      if (previous !== undefined) this.styleChanges.push([[id, previous], -1]);
      const next = (previous ?? -1) + 1;
      this.versions.set(id, next);
      this.styleChanges.push([[id, next], 1]);
    }
  }
  flush() {
    if (!this.edgeChanges.length && !this.styleChanges.length)
      return new Set<string>();
    this.edges.sendData(this.version, new MultiSet(this.edgeChanges));
    this.styles.sendData(this.version, new MultiSet(this.styleChanges));
    this.version++;
    this.edges.sendFrontier(this.version);
    this.styles.sendFrontier(this.version);
    this.graph.run();
    const result = this.affected;
    this.affected = new Set();
    this.edgeChanges = [];
    this.styleChanges = [];
    return result;
  }
}
