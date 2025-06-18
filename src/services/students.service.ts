import { AppDataSource } from "../data-source";
import { Student } from "../entities/Students";
import * as xlsx from "node-xlsx";
import fs from "fs";
import certificateService from "./certificate.service";
import { VerifyIfStudentIsValid } from "../utils/VerifyIfStudentIsValid";
import { verifyIfStudentIsInSameBookOrPage } from "../utils/VerifyIfStudentIsInSameBookOrPage";
import { parseExcelDate } from "../utils/parseExcelDate";
import path from "path";
import { tmpdir } from "os";
import { format } from "date-fns";

const cleanRows = (rows: string[][]): string[][] => {
  return rows.filter((row) => {
    const [name] = row;

    // Ignora se for a linha de header duplicada
    if (name?.trim() === "Nome do Aluno") return false;

    return true;
  });
};

class StudentService {
  async getAll(): Promise<Student[]> {
    const studentRepo = AppDataSource.getRepository(Student);
    const students = await studentRepo.find({
      relations: { certificate: true },
      order: { name: "ASC" },
    });
    return students;
  }

  async importFromXlsx(filePath: string): Promise<{
    errors: string[];
    successCount: number;
  }> {
    if (!fs.existsSync(filePath)) {
      throw new Error(`Arquivo não encontrado: ${filePath}`);
    }

    const workSheets = xlsx.parse(fs.readFileSync(filePath));
    let rows = workSheets[0].data;

    console.log("Antes do filtro:", rows.slice(0, 10));

    rows = cleanRows(rows);

    console.log("Depois do filtro:", rows.slice(0, 10));

    const studentRepo = AppDataSource.getRepository(Student);
    const newStudents: { name: string; registration: string; row: any[] }[] =
      [];

    for (let i = 2; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.length === 0) continue;

      const [name, registration] = row;
      if (!name) continue; // nome é obrigatório

      // Aceita matrícula vazia, mas garante que será string
      newStudents.push({
        name: String(name).trim(),
        registration: registration ? String(registration).trim() : "",
        row,
      });
    }

    const existingRecords = await studentRepo.find({
      relations: { certificate: true },
    });

    const alreadyRegistered = existingRecords.map((student) => ({
      name: student.name.trim(),
      registration: student.registration?.trim() || "",
    }));

    const { validStudents, errors: duplicateErrors } = VerifyIfStudentIsValid(
      newStudents,
      alreadyRegistered
    );

    const preparedExisting = existingRecords.map((student) => ({
      name: student.name.trim(),
      registration: student.registration || "",
      book: student.certificate?.book?.trim() ?? "",
      book_page: student.certificate?.book_page?.trim() ?? "",
    }));

    const preparedNew = newStudents.map(({ name, registration, row }) => ({
      name: name.trim(),
      registration: registration || "",
      book: String(row[6] || "").trim(),
      book_page: String(row[7] || "").trim(),
    }));

    const { errors: bookPageErrors, invalidKeys: bookPageInvalidKeys } =
      verifyIfStudentIsInSameBookOrPage(preparedNew, preparedExisting);

    const allErrors = [...duplicateErrors, ...bookPageErrors];

    // Filtra alunos válidos que não estão em conflito de livro/página
    const filteredValidStudents = validStudents.filter(
      ({ name, registration }) => {
        const key = `${name}|${registration}`;
        return !bookPageInvalidKeys.has(key);
      }
    );

    let successCount = 0;

    const safeDate = (date: any) => {
      const parsed = parseExcelDate(date);
      return parsed instanceof Date ? parsed : null;
    };

    for (const studentData of filteredValidStudents) {
      const [
        _name,
        _registration,
        publication_date,
        publication_page,
        certificate_number,
        second_issue,
        book,
        book_page,
        enrollment_start,
        enrollment_end,
        process_number,
      ] = studentData.row;

      try {
        const student = studentRepo.create({
          name: String(_name).trim(),
          registration: _registration ? String(_registration).trim() : "",
        });

        const certificate = await certificateService.createCertificateFromRow({
          publication_date: safeDate(publication_date),
          publication_page,
          certificate_number,
          second_issue,
          book,
          book_page,
          enrollment_start: safeDate(enrollment_start),
          enrollment_end: safeDate(enrollment_end),
          process_number,
        });

        student.certificate = certificate;
        await studentRepo.save(student);
        successCount++;
      } catch (error) {
        console.error(`Erro ao salvar aluno ${_name}:`, error);
        allErrors.push(`Erro ao processar aluno ${_name}: ${error.message}`);
      }
    }

    fs.unlinkSync(filePath);

    return { errors: allErrors, successCount };
  }

  async exportToXlsx(): Promise<string> {
    const studentRepo = AppDataSource.getRepository(Student);
    const students = await studentRepo.find({
      relations: { certificate: true },
      order: { name: "ASC" },
    });

    // Duas linhas de cabeçalho, com merges para os grupos
    const headerRow1 = [
      "Nome do Aluno",
      "Matrícula",
      "Diário Oficial",
      null,
      "Número",
      null,
      "CECIERJ/CEJA",
      null,
      "DATAS - MATRÍCULAS",
      null,
      "Nº DO PROCESSO",
    ];

    const headerRow2 = [
      null,
      null,
      "Data",
      "Página",
      "1ª Via",
      "2ª Via",
      "Livro",
      "Página",
      "Início",
      "Término",
      null,
    ];

    // Função auxiliar para formatar data no formato dd/MM/yyyy
    const formatDate = (value?: Date | string): string => {
      if (!value) return "";
      const date = new Date(value);
      if (isNaN(date.getTime())) return "";
      return format(date, "dd/MM/yyyy");
    };

    const dataRows = students.map((student) => {
      const cert = student.certificate;
      return [
        student.name || "",
        student.registration || "",
        formatDate(cert?.publication_date),
        cert?.publication_page || "",
        cert?.certificate_number || "",
        cert?.second_issue || "",
        cert?.book || "",
        cert?.book_page || "",
        formatDate(cert?.enrollment_start),
        formatDate(cert?.enrollment_end),
        cert?.process_number || "",
      ];
    });

    const worksheetData = [headerRow1, headerRow2, ...dataRows];

    // Definindo merges para os cabeçalhos agrupados
    const merges = [
      { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } }, // Nome do Aluno
      { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } }, // Matrícula
      { s: { r: 0, c: 10 }, e: { r: 1, c: 10 } }, // Nº DO PROCESSO

      { s: { r: 0, c: 2 }, e: { r: 0, c: 3 } }, // Diário Oficial (Data + Página)
      { s: { r: 0, c: 4 }, e: { r: 0, c: 5 } }, // Número (1ª Via + 2ª Via)
      { s: { r: 0, c: 6 }, e: { r: 0, c: 7 } }, // CECIERJ/CEJA (Livro + Página)
      { s: { r: 0, c: 8 }, e: { r: 0, c: 9 } }, // DATAS - MATRÍCULAS (Início + Término)
    ];

    const buffer = xlsx.build([
      {
        name: "Alunos",
        data: worksheetData,
        options: {
          "!merges": merges,
        },
      },
    ]);

    const fileName = `students_export_${Date.now()}.xlsx`;
    const filePath = path.join(tmpdir(), fileName);

    fs.writeFileSync(filePath, buffer);
    return filePath;
  }
}

export default new StudentService();
