import { MigrationInterface, QueryRunner, Table } from "typeorm";

export class CreateStudents1747365198585 implements MigrationInterface {
  public async up(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.createTable(
      new Table({
        name: "students",
        columns: [
          {
            name: "id",
            type: "uuid",
            isPrimary: true,
            generationStrategy: "uuid",
            default: "uuid_generate_v4()",
          },
          {
            name: "name",
            type: "varchar",
            isNullable: false,
          },
          {
            name: "registration",
            type: "varchar",
            isNullable: true,
          },
          {
            name: "certificateId",
            type: "uuid",
            isNullable: true,
          },
        ],
        foreignKeys: [
          {
            columnNames: ["certificateId"],
            referencedTableName: "certificates",
            referencedColumnNames: ["id"],
            onDelete: "CASCADE",
          },
        ],
      }),
      true
    );
  }

  public async down(queryRunner: QueryRunner): Promise<void> {
    await queryRunner.dropTable("students");
  }
}
