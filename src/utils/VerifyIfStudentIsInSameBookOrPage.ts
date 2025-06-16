import { StudentType } from "../types/student-type";

type ExtendedStudentInput = Omit<StudentType, "certificate"> & {
  book: string;
  book_page: string;
};

type BookPageKey = `${string}|${string}`;
export const verifyIfStudentIsInSameBookOrPage = (
  newStudents: ExtendedStudentInput[],
  existingStudents: ExtendedStudentInput[]
): { errors: string[]; invalidKeys: Set<string> } => {
  const errors: string[] = [];
  const invalidKeys = new Set<string>();

  const existingByName = new Map<string, Set<BookPageKey>>();
  const existingBookPageMap = new Map<BookPageKey, Set<string>>();

  for (const student of existingStudents) {
    const key = `${student.book}|${student.book_page}` as BookPageKey;

    if (!existingByName.has(student.name)) {
      existingByName.set(student.name, new Set());
    }
    existingByName.get(student.name)!.add(key);

    if (!existingBookPageMap.has(key)) {
      existingBookPageMap.set(key, new Set());
    }
    existingBookPageMap.get(key)!.add(student.name);
  }

  const seenNewByBookPage = new Map<BookPageKey, Set<string>>();

  for (const student of newStudents) {
    const { name, registration, book, book_page } = student;
    const key = `${book}|${book_page}` as BookPageKey;
    const fullKey = `${name}|${registration}`;

    // REGRA 1: Mesmo aluno não pode estar no mesmo livro/página
    if (existingByName.get(name)?.size > 0) {
      errors.push(
        `O aluno "${name}" já está registrado em outro livro/página.`
      );
      invalidKeys.add(fullKey);
      continue;
    }

    // REGRA 2: Página não pode ser compartilhada
    const existingNamesInPage = existingBookPageMap.get(key) ?? new Set();
    if (existingNamesInPage.size > 0) {
      errors.push(
        `A página "${book_page}" do livro "${book}" já está ocupada por outro aluno no banco de dados: "${[
          ...existingNamesInPage,
        ].join('", "')}". "${name}" não pode ser registrado nela.`
      );
      invalidKeys.add(fullKey);
      continue;
    }

    // Verificar conflitos na própria planilha
    const newNamesInPage = seenNewByBookPage.get(key) ?? new Set();
    if (newNamesInPage.size > 0) {
      errors.push(
        `A página "${book_page}" do livro "${book}" já está sendo usada por outro aluno na planilha: "${[
          ...newNamesInPage,
        ].join('", "')}". "${name}" não pode ser registrado nela.`
      );
      invalidKeys.add(fullKey);
      continue;
    }

    // Registrar como válido
    if (!seenNewByBookPage.has(key)) {
      seenNewByBookPage.set(key, new Set());
    }
    seenNewByBookPage.get(key)!.add(name);
  }

  return { errors, invalidKeys };
};
