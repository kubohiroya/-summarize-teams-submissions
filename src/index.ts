import fs from "fs";
import parse from "csv-parse";
import XLSX from "xlsx";

const ASSIGNMENT_TITLES = [
  "2020年07月 _memorandum_の提出 (2020年度1Q2Q期間について「1年",
  "2020年8月・9月 _memorandum_の提出 (「1年次」シートに欄に記入)",
  "2020年10月 _memorandum_の提出 (「1年次」シートに欄に記入)",
  "2020年11月 _memorandum_の提出 (「1年次」シートに欄に記入)",
  "2020年12月 _memorandum_の提出 (「1年次」シートに欄に記入)",
  "2021年1月 _memorandum_の提出 (「1年次」シートに欄に記入)",
];

const getHeaderColumns = () => {
  const assignmentMonths = ["7月", "8・9月", "10月", "11月", "12月", "1月"];
  const reportedMonths = ["7月", "8月", "9月", "10月", "11月", "12月", "1月"];
  const keepProblemTry = ["K", "P", "T"];
  return [
    "name",
    "email",
    ...assignmentMonths.map((r) => "提出?" + r),
    ...reportedMonths.map((m) => "学修:幸福度" + m),
    ...reportedMonths.map((m) => "生活:幸福度" + m),
    "(文字数)学修:目標",
    "(文字数)生活:目標",
    ...reportedMonths
      .map((r) => keepProblemTry.map((i) => "(文字数)学修:" + i + r))
      .flat(),
    ...reportedMonths
      .map((r) => keepProblemTry.map((i) => "(文字数)生活:" + i + r))
      .flat(),
    "学習:目標",
    "生活:目標",
    ...reportedMonths
      .map((r) => keepProblemTry.map((i) => "学修:" + i + r))
      .flat(),
    ...reportedMonths
      .map((r) => keepProblemTry.map((i) => "生活:" + i + r))
      .flat(),
  ];
};

const HEADER_COLUMNS = getHeaderColumns();
const SUBMISSION_VERSION_TAG = 'バージョン 1';

const getSrcItem = (excelFile: string | undefined) => {
  const sheetName: string = "1年次";
  const columns = [8, 9, 10, 11, 12, 13, 14];
  if (excelFile) {
    const workbook = XLSX.readFile(excelFile);
    const sheet = workbook.Sheets[sheetName];
    return {
      studyGoal: sheet["B2"]?.w || "",
      otherGoal: sheet["F2"]?.w || "",
      study: columns.map((col) => ({
        k: sheet["B" + col]?.w || "",
        p: sheet["C" + col]?.w || "",
        t: sheet["D" + col]?.w || "",
        h: sheet["E" + col]?.v || null,
      })),
      other: columns.map((col) => ({
        k: sheet["F" + col]?.w || "",
        p: sheet["G" + col]?.w || "",
        t: sheet["H" + col]?.w || "",
        h: sheet["I" + col]?.v || null,
      })),
    };
  } else {
    return {
      studyGoal: "",
      otherGoal: "",
      study: [...Array(columns.length)].map(() => ({
        k: "",
        p: "",
        t: "",
        h: null,
      })),
      other: [...Array(columns.length)].map(() => ({
        k: "",
        p: "",
        t: "",
        h: null,
      })),
    };
  }
};

const convertSrcItemToRow = (
  name: string,
  email: string,
  hasSubmit: number[],
  src: any
) => {
  type KPT = {
    k: string;
    p: string;
    t: string;
    h: number;
  };
  return [
    name,
    email,
    ...hasSubmit,
    ...src.study.map((m: KPT) => m.h).flat(),
    ...src.other.map((m: KPT) => m.h).flat(),
    src.studyGoal.length || 0,
    src.otherGoal.length || 0,
    ...src.study.map((m: KPT) => [m.k.length, m.p.length, m.t.length]),
    ...src.other.map((m: KPT) => [m.k.length, m.p.length, m.t.length]),
    src.studyGoal,
    src.otherGoal,
    ...src.study.map((m: KPT) => [m.k, m.p, m.t]),
    ...src.other.map((m: KPT) => [m.k, m.p, m.t]),
  ].flat();
};

const main = async (
  inputFile: string,
  unzipDir: string,
  outputFile: string,
  sheetName: string
) => {
  const getStudents = async (filename: string) => {
    const users = [];
    const parser = fs
      .createReadStream(filename)
      .pipe(parse({ delimiter: "\t" }));
    for await (const record of parser) {
      users.push(record);
    }
    return users;
  };

  const students = await getStudents(inputFile);

  const items = await Promise.all(
    students.map(async (student) => {
      const [name, email] = student;

      const files: Array<undefined | string> = await Promise.all(
        ASSIGNMENT_TITLES.map(async (assignmentName) => {
          try {
            const basePath = `${unzipDir}/${name}/${assignmentName}/${SUBMISSION_VERSION_TAG}`;
            const dir = await fs.readdirSync(basePath);
            return `${basePath}/${dir[0]}`;
          } catch (ignore) {
            return undefined;
          }
        })
      );

      const hasSubmit = files.map((file) => (file ? 1 : 0));
      const lastFile = files.filter((file) => file).pop();

      return convertSrcItemToRow(name, email, hasSubmit, getSrcItem(lastFile));
    })
  );

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([HEADER_COLUMNS].concat(items));
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, outputFile);
};

if (process.argv.length != 6) {
  console.error(
    "Usage: ts-node index.ts MEMBER.csv UnzippedDir OUTPUT.xlsb, SheetName"
  );
  console.error(
    "eg. ts-node index.ts ./res/member.csv ./res/提出済みのファイル output.xlsb 2021"
  );
  process.exit(-1);
}

main(process.argv[2], process.argv[3], process.argv[4], process.argv[5]);
