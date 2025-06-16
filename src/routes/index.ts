import { Express } from "express";
import studentRouter from "./student.route";

const Routers = (app: Express): void => {
  app.use("/students", studentRouter);
};

export default Routers;
