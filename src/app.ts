import express from "express";
import Routers from "./routes";
import cors from "cors";

const app = express();

app.use(express.json());

app.use(cors());

Routers(app);

export default app;
