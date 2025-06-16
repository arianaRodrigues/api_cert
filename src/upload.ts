import multer from "multer";
import path from "path";
import crypto from "crypto";

const tmpFolder = path.resolve(__dirname, "..", "..", "tmp");

const storage = multer.diskStorage({
  destination: tmpFolder,
  filename: (req, file, cb) => {
    const hash = crypto.randomBytes(16).toString("hex");
    const filename = `${hash}-${file.originalname}`;
    cb(null, filename);
  },
});

export const upload = multer({ storage });
