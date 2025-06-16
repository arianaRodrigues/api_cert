import app from "./app";
import { AppDataSource } from "./data-source";

AppDataSource.initialize()
  .then(() => {
    const port = process.env.PORT || 3333;

    app.listen(port, () => {
      console.log(`App running on http://localhost:${port}/`);
    });
  })
  .catch((err) => console.error(err));
