import { Router } from "express";
import { upload } from "../upload";
import studentController from "../controllers/student.controller";

const studentRouter = Router();

studentRouter.get("/", studentController.getAllStudents);
studentRouter.post(
  "/upload",
  upload.single("file"),
  studentController.importFromXlsx
);
studentRouter.get("/export", studentController.exportToXlsx);

export default studentRouter;
