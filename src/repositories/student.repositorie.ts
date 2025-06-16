import { DeleteResult, Repository, UpdateResult } from "typeorm";
import { Student } from "../entities/Students";
import { AppDataSource } from "../data-source";

interface IStudentRepo {
  save: (user: Partial<Student>) => Promise<Student>;
  all: () => Promise<Student[]>;
  findOne: (payload: object) => Promise<Student>;
  findById: (id: string) => Promise<Student>;
  update: (id: string, payload: Partial<Student>) => Promise<UpdateResult>;
  deleteById: (id: string) => Promise<DeleteResult>;
}

class StudentRepository implements IStudentRepo {
  private ormRepo: Repository<Student>;

  constructor() {
    this.ormRepo = AppDataSource.getRepository(Student);
  }
  findById = async (id: string) => {
    return await this.ormRepo.findOne({ where: { id: id } });
  };
  findOne = async (payload: object) => {
    return await this.ormRepo.findOne({ ...payload });
  };
  save = async (user: Partial<Student>) => await this.ormRepo.save(user);

  all = async () => await this.ormRepo.find();

  update = async (id: string, payload: Partial<Student>) =>
    await this.ormRepo.update(id, { ...payload });

  deleteById = async (id: string) => {
    return await this.ormRepo.delete(id);
  };
}

export default new StudentRepository();
