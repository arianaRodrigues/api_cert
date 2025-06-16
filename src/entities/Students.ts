import {
  Entity,
  PrimaryGeneratedColumn,
  Column,
  OneToOne,
  JoinColumn,
} from "typeorm";
import { Certificate } from "./Certificates";

@Entity("students")
export class Student {
  @PrimaryGeneratedColumn("uuid")
  id?: string;

  @Column()
  name: string;

  @Column({ nullable: true })
  registration: string;

  @OneToOne(() => Certificate, (certificate) => certificate.student, {
    cascade: true,
    onDelete: "CASCADE",
  })
  @JoinColumn()
  certificate: Certificate;
}
