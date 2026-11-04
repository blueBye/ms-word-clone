import { readFile, readdir } from "node:fs/promises";
import { dirname, resolve, relative } from "node:path";
import { expect, it } from "vitest";
import ts from "typescript";

it("keeps the core import graph independent of OOXML, React, Next.js and browser DOM types", async () => {
  const visited = new Set<string>();
  const root = resolve("src");
  const visit = async (file: string): Promise<void> => {
    if (visited.has(file)) return;
    visited.add(file);
    expect(relative(root, file).replaceAll("\\", "/")).toMatch(
      /^(core|model)\//,
    );
    const ast = ts.createSourceFile(
      file,
      await readFile(file, "utf8"),
      ts.ScriptTarget.Latest,
      true,
    );
    const imports: string[] = [];
    const inspect = (node: ts.Node) => {
      if (
        (ts.isImportDeclaration(node) || ts.isExportDeclaration(node)) &&
        node.moduleSpecifier &&
        ts.isStringLiteral(node.moduleSpecifier)
      )
        imports.push(node.moduleSpecifier.text);
      if (ts.isIdentifier(node))
        expect(
          [
            "window",
            "HTMLElement",
            "DOMParser",
            "XMLSerializer",
            "React",
          ].includes(node.text),
        ).toBe(false);
      ts.forEachChild(node, inspect);
    };
    inspect(ast);
    for (const specifier of imports) {
      if (specifier.startsWith("."))
        await visit(resolve(dirname(file), specifier).replace(/\.js$/, ".ts"));
      else expect(specifier).toBe("@electric-sql/d2ts");
    }
  };
  await visit(resolve(root, "core/index.ts"));
  expect(visited.size).toBeGreaterThan(10);
  for (const name of await readdir("apps/web/components/editor"))
    if (/\.(ts|tsx)$/.test(name))
      expect(
        await readFile(`apps/web/components/editor/${name}`, "utf8"),
      ).not.toContain("@local/docx-ir-engine/docx");
});
