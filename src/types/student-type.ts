export type StudentType = {
  id?: string;
  name: string;
  registration: string | null;
  certificate: {
    id?: string;
    publication_date: Date | null;
    publication_page: string | null;
    certificate_number: string | null;
    second_issue: string | null;
    book: string | null;
    book_page: string | null;
    enrollment_start: Date | null;
    enrollment_end: Date | null;
    process_number: string | null;
  };
};
