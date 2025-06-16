import { AppDataSource } from "../data-source";
import { Certificate } from "../entities/Certificates";

type CertificateRowData = {
  publication_date: any;
  publication_page: any;
  certificate_number: any;
  second_issue: any;
  book: any;
  book_page: any;
  enrollment_start: any;
  enrollment_end: any;
  process_number: any;
};

class CertificateService {
  createCertificateFromRow(data: CertificateRowData): Certificate {
    const certificateRepo = AppDataSource.getRepository(Certificate);

    const certificate = certificateRepo.create({
      publication_date: data.publication_date
        ? new Date(data.publication_date)
        : null,
      publication_page: String(data.publication_page || ""),
      certificate_number: String(data.certificate_number || ""),
      second_issue: String(data.second_issue || ""),
      book: String(data.book || ""),
      book_page: String(data.book_page || ""),
      enrollment_start: data.enrollment_start
        ? new Date(data.enrollment_start)
        : null,
      enrollment_end: data.enrollment_end
        ? new Date(data.enrollment_end)
        : null,
      process_number: String(data.process_number || ""),
    });

    return certificate;
  }
}

export default new CertificateService();
