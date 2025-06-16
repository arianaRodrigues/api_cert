import { Entity, Column, OneToOne, PrimaryGeneratedColumn } from "typeorm";
import { Student } from "./Students";

@Entity("certificates")
export class Certificate {
  @PrimaryGeneratedColumn("uuid")
  id?: string;

  @OneToOne(() => Student, (student) => student.certificate)
  student: Student;

  @Column({ type: "date", nullable: true })
  publication_date: Date;

  @Column({ nullable: true })
  publication_page: string;

  @Column({ nullable: true })
  certificate_number: string;

  @Column({ nullable: true })
  second_issue: string;

  @Column({ nullable: true })
  book: string;

  @Column({ nullable: true })
  book_page: string;

  @Column({ type: "date", nullable: true })
  enrollment_start: Date;

  @Column({ type: "date", nullable: true })
  enrollment_end: Date;

  @Column({ nullable: true })
  process_number: string;
}
