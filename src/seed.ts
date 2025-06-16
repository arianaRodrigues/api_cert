import { DataSource } from "typeorm";
import { Certificate } from "./entities/Certificates";
import { Student } from "./entities/Students";

export async function seedDatabase(dataSource: DataSource) {
  await dataSource.initialize();

  // Criar certificados fictícios
  const certificate1 = dataSource.manager.create(Certificate, {
    id: "a1111111-1111-1111-1111-111111111111",
    publication_date: new Date("2023-01-15"),
    publication_page: "15",
    certificate_number: "CERT-2023-0001",
    second_issue: "No",
    book: "Livro 1",
    book_page: "100",
    enrollment_start: new Date("2022-01-01"),
    enrollment_end: new Date("2022-12-31"),
    process_number: "PROC-12345",
  });

  const certificate2 = dataSource.manager.create(Certificate, {
    id: "b2222222-2222-2222-2222-222222222222",
    publication_date: new Date("2024-02-20"),
    publication_page: "42",
    certificate_number: "CERT-2024-0002",
    second_issue: "Yes",
    book: "Livro 2",
    book_page: "200",
    enrollment_start: new Date("2023-02-01"),
    enrollment_end: new Date("2023-11-30"),
    process_number: "PROC-54321",
  });

  await dataSource.manager.save([certificate1, certificate2]);

  // Criar alunos fictícios vinculados aos certificados
  const student1 = dataSource.manager.create(Student, {
    id: "c3333333-3333-3333-3333-333333333333",
    name: "João Silva",
    registration: "REG-1001",
    certificateId: certificate1.id,
  });

  const student2 = dataSource.manager.create(Student, {
    id: "d4444444-4444-4444-4444-444444444444",
    name: "Maria Oliveira",
    registration: "REG-1002",
    certificateId: certificate2.id,
  });

  await dataSource.manager.save([student1, student2]);

  console.log("Seed concluído com sucesso!");
  await dataSource.destroy();
}
