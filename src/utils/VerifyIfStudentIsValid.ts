type StudentInput = {
  name: string;
  registration: string;
  row?: any[];
};

const getSegment = (
  registration: string
): "elementary" | "highschool" | "unknown" => {
  const segment = registration.substring(2, 4);
  if (segment === "10") return "elementary";
  if (segment === "20") return "highschool";
  return "unknown";
};

export const VerifyIfStudentIsValid = (
  newStudents: StudentInput[],
  existingStudents: StudentInput[]
): {
  validStudents: StudentInput[];
  errors: string[];
} => {
  const validStudents: StudentInput[] = [];
  const errors: string[] = [];

  const existingMap = new Map<string, Set<string>>();

  for (const student of existingStudents) {
    const segment = getSegment(student.registration);
    if (!existingMap.has(student.name)) {
      existingMap.set(student.name, new Set());
    }
    existingMap.get(student.name)?.add(segment);
  }

  for (const student of newStudents) {
    const segment = getSegment(student.registration);
    const existingSegments = existingMap.get(student.name);

    if (existingSegments?.has(segment)) {
      errors.push(
        `O Aluno "${student.name}" com a matrícula "${student.registration}" já existe no sistema.`
      );
    } else {
      validStudents.push(student);

      if (!existingMap.has(student.name)) {
        existingMap.set(student.name, new Set());
      }
      existingMap.get(student.name)?.add(segment);
    }
  }

  return { validStudents, errors };
};
