/** Shared, code-owned vector artwork. Coordinates use a 100 by 100 canvas. */
export const shapePaths: Record<string, string> = {
  triangle: "M50 0 100 100H0Z",
  rtTriangle: "M0 0 100 100H0Z",
  diamond: "M50 0 100 50 50 100 0 50Z",
  hexagon: "M25 0H75L100 50 75 100H25L0 50Z",
  pentagon: "M50 0 100 38 80 100H20L0 38Z",
  chevron: "M0 0H60L100 50 60 100H0L40 50Z",
  rightArrow: "M0 30H60V0L100 50 60 100V70H0Z",
  leftArrow: "M100 30H40V0L0 50 40 100V70H100Z",
  upArrow: "M30 100V40H0L50 0 100 40H70V100Z",
  downArrow: "M30 0V60H0L50 100 100 60H70V0Z",
  star5: "M50 0 61 35 98 35 68 57 79 94 50 72 21 94 32 57 2 35 39 35Z",
  plus: "M35 0H65V35H100V65H65V100H35V65H0V35H35Z",
  heart: "M50 95 7 52C-15 22 15-15 50 20 85-15 115 22 93 52Z",
  line: "M0 100 100 0",
};
export const iconLibrary = [
  {
    id: "home",
    name: "Home",
    category: "Buildings",
    path: "M5 45L50 5L95 45L85 55L80 50L80 95L60 95L60 65L40 65L40 95L20 95L20 50L15 55Z",
  },
  {
    id: "building",
    name: "Office building",
    category: "Buildings",
    path: "M15 5H85V95H60V75H40V95H15ZM25 20V30H35V20ZM45 20V30H55V20ZM65 20V30H75V20ZM25 40V50H35V40ZM45 40V50H55V40ZM65 40V50H75V40Z",
  },
  {
    id: "person",
    name: "Person",
    category: "People",
    path: "M50 5C28 5 28 37 50 37C72 37 72 5 50 5ZM25 45H75L90 90H10Z",
  },
  {
    id: "people",
    name: "Team",
    category: "People",
    path: "M35 8C17 8 17 34 35 34C53 34 53 8 35 8ZM70 18C55 18 55 40 70 40C85 40 85 18 70 18ZM15 42H55L68 90H2ZM62 48H84L98 90H76Z",
  },
  {
    id: "mail",
    name: "Mail",
    category: "Communication",
    path: "M5 15H95V85H5ZM15 25L50 52L85 25H75L50 43L25 25ZM15 38V75H85V38L50 65Z",
  },
  {
    id: "chat",
    name: "Conversation",
    category: "Communication",
    path: "M5 10H95V70H40L15 95V70H5ZM20 25V35H80V25ZM20 45V55H65V45Z",
  },
  {
    id: "phone",
    name: "Phone",
    category: "Communication",
    path: "M15 5H35L43 28L30 40C38 56 45 64 60 70L73 57L95 65V85C95 110 5 65 5 20Z",
  },
  {
    id: "calendar",
    name: "Calendar",
    category: "Business",
    path: "M20 5H30V15H70V5H80V15H95V95H5V15H20ZM15 35V85H85V35ZM25 45V55H40V45ZM50 45V55H65V45ZM25 65V75H40V65ZM50 65V75H65V65Z",
  },
  {
    id: "briefcase",
    name: "Briefcase",
    category: "Business",
    path: "M30 5H70V25H95V90H5V25H30ZM40 15V25H60V15ZM15 40V52H45V45H55V52H85V40ZM15 62V80H85V62H55V70H45V62Z",
  },
  {
    id: "chart",
    name: "Analytics",
    category: "Business",
    path: "M5 5H15V85H95V95H5ZM25 55H40V75H25ZM50 35H65V75H50ZM75 15H90V75H75Z",
  },
  {
    id: "computer",
    name: "Computer",
    category: "Technology",
    path: "M5 10H95V75H55V85H75V95H25V85H45V75H5ZM15 20V65H85V20Z",
  },
  {
    id: "database",
    name: "Database",
    category: "Technology",
    path: "M10 20C10 0 90 0 90 20V80C90 105 10 105 10 80ZM20 20C20 30 80 30 80 20C80 10 20 10 20 20ZM20 38V48C35 60 65 60 80 48V38C65 48 35 48 20 38ZM20 65V78C35 90 65 90 80 78V65C65 75 35 75 20 65Z",
  },
  {
    id: "lock",
    name: "Lock",
    category: "Technology",
    path: "M25 40V25C25 -5 75 -5 75 25V40H90V95H10V40ZM35 40H65V25C65 8 35 8 35 25ZM45 58V78H55V58Z",
  },
  {
    id: "shield",
    name: "Security shield",
    category: "Technology",
    path: "M50 2L92 18V48C92 73 65 92 50 98C35 92 8 73 8 48V18ZM24 48L44 68L78 34L70 26L44 52L32 40Z",
  },
  {
    id: "book",
    name: "Book",
    category: "Education",
    path: "M5 10C20 5 35 5 50 15C65 5 80 5 95 10V85C80 80 65 80 50 90C35 80 20 80 5 85ZM15 20V73C25 70 35 72 45 77V25C35 18 25 17 15 20ZM55 25V77C65 72 75 70 85 73V20C75 17 65 18 55 25Z",
  },
  {
    id: "cap",
    name: "Graduation",
    category: "Education",
    path: "M0 30L50 5L100 30L50 55ZM20 50L50 65L80 50V75C65 95 35 95 20 75ZM90 40H98V85H90Z",
  },
  {
    id: "check",
    name: "Check mark",
    category: "Signs",
    path: "M5 50L20 35L40 55L80 10L95 25L40 90Z",
  },
  {
    id: "warning",
    name: "Warning",
    category: "Signs",
    path: "M50 2L100 95H0ZM44 30V65H56V30ZM44 75V85H56V75Z",
  },
  {
    id: "pin",
    name: "Location",
    category: "Travel",
    path: "M50 2C5 2 2 45 20 65L50 98L80 65C98 45 95 2 50 2ZM50 20C30 20 30 50 50 50C70 50 70 20 50 20Z",
  },
  {
    id: "plane",
    name: "Airplane",
    category: "Travel",
    path: "M43 5H57V35L95 60V73L57 58V82L70 92V100L50 94L30 100V92L43 82V58L5 73V60L43 35Z",
  },
  {
    id: "leaf",
    name: "Leaf",
    category: "Nature",
    path: "M90 5C95 70 70 100 20 85L5 100L0 92L15 77C0 30 35 5 90 5ZM25 70L70 25L62 22L22 62Z",
  },
  {
    id: "music",
    name: "Music",
    category: "Arts",
    path: "M35 15L90 5V75C90 100 55 100 55 80C55 68 70 64 80 68V25L45 32V85C45 105 10 105 10 85C10 73 25 69 35 73Z",
  },
  {
    id: "camera",
    name: "Camera",
    category: "Arts",
    path: "M30 10H70L80 25H95V90H5V25H20ZM50 35C20 35 20 80 50 80C80 80 80 35 50 35ZM50 45C67 45 67 70 50 70C33 70 33 45 50 45Z",
  },
  {
    id: "folder",
    name: "Folder",
    category: "Business",
    path: "M5 15H40L50 28H95V90H5ZM15 40V80H85V40Z",
  },
] as const;
export const iconPaths: Record<string, string> = Object.fromEntries(
  iconLibrary.map((icon) => [`icon_${icon.id}`, icon.path]),
);

/** Convert the small SVG path grammar used by the trusted catalogue to absolute drawing commands. */
export function vectorCommands(
  path: string,
): { kind: "M" | "L" | "C" | "Q" | "Z"; points: number[] }[] {
  const tokens = path.match(/[MLCQHVZ]|-?(?:\d*\.)?\d+/g) ?? [],
    result: ReturnType<typeof vectorCommands> = [];
  let i = 0,
    command = "",
    x = 0,
    y = 0;
  while (i < tokens.length) {
    if (/^[A-Z]$/.test(tokens[i])) command = tokens[i++];
    if (command === "Z") {
      result.push({ kind: "Z", points: [] });
      command = "";
      continue;
    }
    const count =
      command === "C"
        ? 6
        : command === "Q"
          ? 4
          : command === "H" || command === "V"
            ? 1
            : 2;
    if (!command || i + count > tokens.length)
      throw new Error("Invalid catalogue path.");
    const points = tokens.slice(i, i + count).map(Number);
    i += count;
    if (points.some((n) => !Number.isFinite(n)))
      throw new Error("Invalid catalogue coordinates.");
    if (command === "H") {
      x = points[0];
      result.push({ kind: "L", points: [x, y] });
    } else if (command === "V") {
      y = points[0];
      result.push({ kind: "L", points: [x, y] });
    } else {
      x = points.at(-2)!;
      y = points.at(-1)!;
      result.push({ kind: command as "M" | "L" | "C" | "Q", points });
    }
    if (command === "M") command = "L";
  }
  return result;
}
