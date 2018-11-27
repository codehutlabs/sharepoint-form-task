import { ICompanyItem } from '../models/ICompanyItem';
import { IRobertKuzmaItem } from '../models/IRobertKuzmaItem';

export interface ISharepointCandidatesState {
  debug: boolean;
  companies: ICompanyItem[];
  items: IRobertKuzmaItem[];
  chartData: {};
  formValid: boolean;
  values: {
    lastName: string;
    firstName: string;
    email: string;
    company: string;
    salary: string;
  };
  touched: {
    lastName?: boolean;
    firstName?: boolean;
    email?: boolean;
    company?: boolean;
    salary?: boolean;
  };
  errors: {
    lastName?: boolean;
    firstName?: boolean;
    email?: boolean;
    company?: boolean;
    salary?: boolean;
  };

}
