export interface Student {
  sbd: string;
  name: string;
  dob: any;
  gender: string;
  lop: string;
  scores: Record<string, number[]>;
}

export const SUBJECTS = ['Anh', 'Địa', 'Hoá', 'Lí', 'Toán', 'Văn', 'Sinh', 'Sử', 'GDCD', 'Tin', 'CN', 'TD'];
