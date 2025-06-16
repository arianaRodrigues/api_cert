import express from "express";
import studentService from "../services/students.service";
import fs from "fs";
class StudentController {
  getAllStudents = async (
    req: express.Request,
    res: express.Response
  ): Promise<express.Response> => {
    try {
      const students = await studentService.getAll();
      return res.status(200).json(students);
    } catch (error) {
      return res.status(500).json({ message: error.message });
    }
  };

  importFromXlsx = async (req: express.Request, res: express.Response) => {
    try {
      const file = req.file;

      if (!file) {
        res.status(400).send({ error: "Nenhum arquivo foi enviado." });
        return;
      }

      const result = await studentService.importFromXlsx(file.path);

      if (result.errors.length > 0) {
        res.status(207).json({
          message: "Importação concluída com alguns erros",
          successCount: result.successCount,
          errorCount: result.errors.length,
          errors: result.errors,
        });
      } else {
        res.status(200).json({
          message: "Todos os alunos foram importados com sucesso.",
          successCount: result.successCount,
        });
      }
    } catch (err) {
      console.error(err);
      res.status(500).send({ error: "Erro interno ao importar." });
    }
  };

  exportToXlsx = async (req: express.Request, res: express.Response) => {
    try {
      const filePath = await studentService.exportToXlsx();

      res.download(filePath, "alunos.xlsx", (err) => {
        if (err) {
          console.error("Erro ao enviar arquivo:", err);
          res.status(500).send({ error: "Erro ao exportar planilha." });
        }

        // Remove o arquivo temporário depois do download
        fs.unlinkSync(filePath);
      });
    } catch (error) {
      console.error("Erro na exportação:", error);
      res.status(500).send({ error: "Erro ao gerar planilha." });
    }
  };
}

export default new StudentController();
