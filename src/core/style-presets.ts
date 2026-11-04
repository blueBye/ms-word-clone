export const builtins: Record<
  string,
  {
    name: string;
    size?: number;
    bold?: boolean;
    after?: number;
    before?: number;
    outline?: number;
  }
> = {
  Normal: { name: "Normal", size: 12, after: 8 },
  NoSpacing: { name: "No Spacing", after: 0 },
  Heading1: {
    name: "Heading 1",
    size: 20,
    bold: true,
    before: 12,
    after: 6,
    outline: 0,
  },
  Heading2: {
    name: "Heading 2",
    size: 16,
    bold: true,
    before: 10,
    after: 4,
    outline: 1,
  },
  Heading3: {
    name: "Heading 3",
    size: 14,
    bold: true,
    before: 8,
    after: 4,
    outline: 2,
  },
  Title: { name: "Title", size: 28, after: 12 },
  Subtitle: { name: "Subtitle", size: 16, after: 8 },
};
