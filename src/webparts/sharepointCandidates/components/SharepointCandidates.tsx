import * as React from 'react';
import styles from './SharepointCandidates.module.scss';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { Fabric, PrimaryButton } from 'office-ui-fabric-react';
import { ISharepointCandidatesProps } from './ISharepointCandidatesProps';
import { ISharepointCandidatesState } from './ISharepointCandidatesState';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISharePointList } from '../models/ISharePointList';
import { ICompanyItem } from '../models/ICompanyItem';
import { IRobertKuzmaItem } from '../models/IRobertKuzmaItem';
import { Pie } from 'react-chartjs-2';


export default class SharepointCandidates extends React.Component<ISharepointCandidatesProps, ISharepointCandidatesState> {

  private _companyListName: string = "Company";
  private _robertKuzmaListName: string = "RobertKuzma";
  private _listsUrl: string;

  constructor(props) {
    super(props);

    this.state = {
      debug: false,
      companies: [],
      items: [],
      chartData: {},
      formValid: false,
      values: {
        lastName: "",
        firstName: "",
        email: "",
        company: "",
        salary: "0"
      },
      touched: {},
      errors: {}
    };

    this._listsUrl = `${props.dataProvider.webPartContext.pageContext.web.absoluteUrl}/_api/Web/Lists`;
  }

  private _createItem(values): void {

    const requester = this.props.dataProvider.webPartContext.spHttpClient;
    const listItemsUrl: string = `${this._listsUrl}/GetByTitle('${this._robertKuzmaListName}')/Items`;

    const body: string = JSON.stringify({
      'Title': values['lastName'],
      'FirstName': values['firstName'],
      'Email': values['email'],
      'Company': values['company'],
      'Salary': values['salary']
    });

    requester.post(listItemsUrl, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata',
          'odata-version': ''
        },
        body: body
      }).then((response: SPHttpClientResponse) => {
        this._loadRobertKuzma();
      });

  }

  private _loadCompanies(): void {

    const requester = this.props.dataProvider.webPartContext.spHttpClient;
    const queryString: string = `?$select=Id,Title`;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._companyListName}')/Items${queryString}`;

    requester.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json: { value: ICompanyItem[] }) => {
        const companies: ICompanyItem[] = json.value.map((item: ICompanyItem) => {
          return {
            Id: item.Id,
            Title: item.Title
          };
        });

        this.setState({
          companies
        });

      });
  }

  private _loadRobertKuzma(): void {

    const requester = this.props.dataProvider.webPartContext.spHttpClient;
    const queryString: string = `?$select=Id,Title,FirstName,Email,Company,Salary`;
    const queryUrl: string = `${this._listsUrl}/GetByTitle('${this._robertKuzmaListName}')/Items${queryString}`;

    requester.get(queryUrl, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((json: { value: IRobertKuzmaItem[] }) => {
        const items: IRobertKuzmaItem[] = json.value.map((item: IRobertKuzmaItem) => {
          return {
            Id: item.Id,
            Title: item.Title,
            FirstName: item.FirstName,
            Email: item.Email,
            Company: item.Company,
            Salary: item.Salary,
          };
        });

        this.setState({
          items
        });

        const backgroundColor = [];
        const data = [];
        const labels = [];

        for (let item of items) {
          backgroundColor.push(this._randomHex());
          data.push(item["Salary"]);
          labels.push(item["Title"]);
        }

        const chartData = {
          datasets: [{
            label: "Candidates",
            backgroundColor: backgroundColor,
            data: data
          }],
          labels: labels
        };

        this.setState({
          chartData
        });

      });
  }


  private handleChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const target = event.target as HTMLInputElement;
    const value: string = target.value;
    const name: string = target.name;

    this.setState({
      values: {
        ...this.state.values,
        [name]: value
      }
    });
  }

  private handleSelectChange = (event: React.ChangeEvent<HTMLSelectElement>): void => {
    const target = event.target as HTMLSelectElement;
    const value: string = target.value;
    const name: string = target.name;

    this.setState({
      values: {
        ...this.state.values,
        [name]: value
      }
    });
  }

  private handleBlure = (event: React.FocusEvent<HTMLInputElement>): void => {
    const target = event.target as HTMLInputElement;
    const name: string = target.name;

    this.setState({
      touched: {
        ...this.state.touched,
        [name]: true
      }
    });

    this.validateForm();
  }

  private handleSelectBlure = (event: React.FocusEvent<HTMLSelectElement>): void => {
    const target = event.target as HTMLSelectElement;
    const name: string = target.name;

    this.setState({
      touched: {
        ...this.state.touched,
        [name]: true
      }
    });

    this.validateForm();
  }

  // Returns if a value is a string
  private isString(value: string): boolean {
    return typeof value === 'string';
  }

  private isNumber(value: string): boolean {
    return typeof value === 'string' && typeof parseInt(value) === 'number' && isFinite(parseInt(value));
  }

  private isEmail(value: string): boolean {
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(value);
  }

  private validateForm(): void {

    const fields = ['lastName', 'firstName', 'email', 'company', 'salary'];
    const errors = {};

    for (let name of fields) {

      if(['lastName', 'firstName', 'email', 'company'].indexOf(name) != -1) {
        if(!this.isString(this.state.values[name]) || this.isString(this.state.values[name]) && this.state.values[name].length < 1) {
          errors[name] = true;
        } else {
          errors[name] = false;
        }
      }

      if(['email'].indexOf(name) != -1) {
        if(!this.isEmail(this.state.values[name])) {
          errors[name] = true;
        } else {
          errors[name] = false;
        }
      }

      if(['salary'].indexOf(name) != -1) {
        if(!this.isNumber(this.state.values[name]) || this.isNumber(this.state.values[name]) && parseInt(this.state.values[name], 10) < 1) {
          errors[name] = true;
        } else {
          errors[name] = false;
        }
      }

    }

    let formValid = true;
    for (let name of fields) {
      if(errors[name]) {
        formValid = false;
        break;
      }
    }

    this.setState({
      formValid,
      errors
    });

  }

  private handleSubmit = (event: React.FormEvent<HTMLFormElement>): void => {
    event.preventDefault();

    const values = JSON.parse(JSON.stringify(this.state.values));

    this._createItem(values);

    this.setState({
      formValid: false,
      values: {
        lastName: "",
        firstName: "",
        email: "",
        company: "",
        salary: "0"
      },
      touched: {},
      errors: {}
    });

  }

  private _randomHex(): string {
    return '#'+Math.floor(Math.random()*16777215).toString(16);
  }

  private _numberToString(value: number): string {
    return `${value}`;
  }

  public componentDidMount(): void {
    this._loadCompanies();
    this._loadRobertKuzma();
  }

  public render(): React.ReactElement<ISharepointCandidatesProps> {

    const {chartData, items } = this.state;
    const chart = JSON.parse(JSON.stringify(chartData));

    return (
      <Fabric>
      <div className={ styles.sharepointCandidates }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>{escape(this.props.description)} Form Validation</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>

              <form onSubmit={this.handleSubmit}>
                <label className={ styles.label }>
                  Last Name:
                  { (this.state.touched.lastName && this.state.errors.lastName) && (
                    <span className={styles.alert}>This field is required!</span>
                  )}
                  <div className={styles.inputField}>
                    <input
                      name="lastName"
                      type="text"
                      value={this.state.values.lastName}
                      onChange={this.handleChange}
                      onBlur={this.handleBlure} />
                  </div>
                </label>
                <label className={ styles.label }>
                  First Name:
                  { this.state.touched.firstName && this.state.errors.firstName && (
                    <span className={styles.alert}>This field is required!</span>
                  )}
                  <div className={styles.inputField}>
                    <input
                      name="firstName"
                      type="text"
                      value={this.state.values.firstName}
                      onChange={this.handleChange}
                      onBlur={this.handleBlure} />
                  </div>
                </label>
                <label className={ styles.label }>
                  Email:
                  { this.state.touched.email && this.state.errors.email && (
                    <span className={styles.alert}>This field is required and must be in a valid format!</span>
                  )}
                  <div className={styles.inputField}>
                    <input
                      name="email"
                      type="text"
                      value={this.state.values.email}
                      onChange={this.handleChange}
                      onBlur={this.handleBlure} />
                  </div>
                </label>
                <label className={ styles.label }>
                  Company:
                  { this.state.touched.company && this.state.errors.company && (
                    <span className={styles.alert}>This field is required!</span>
                  )}
                  <div className={styles.selectField}>
                    <select
                      name="company"
                      value={this.state.values.company}
                      onChange={this.handleSelectChange}
                      onBlur={this.handleSelectBlure}>
                      <option value="">- Select an option -</option>
                      {this.state.companies.map(company => (
                        <option key={company.Id} value={company.Title}>{company.Title}</option>
                      ))}
                    </select>
                  </div>
                </label>
                <label className={ styles.label }>
                  Salary:
                  { this.state.touched.salary && this.state.errors.salary && (
                    <span className={styles.alert}>This field is required and must be greater than 0!</span>
                  )}
                  <div className={styles.inputField}>
                    <input
                      name="salary"
                      type="number"
                      value={this.state.values.salary}
                      onChange={this.handleChange}
                      onBlur={this.handleBlure} />
                  </div>
                </label>
                <PrimaryButton type="submit" disabled={!this.state.formValid}>Send</PrimaryButton>
              </form>

              <Pie data={chart} options={{legend: { display: false }}} />

              <table className={styles.table}>
                <thead>
                  <tr>
                    <th>Id</th>
                    <th>Last Name</th>
                    <th>First Name</th>
                    <th>Email</th>
                    <th>Company</th>
                    <th>Salary</th>
                  </tr>
                </thead>
                <tbody>
                {items.map(item => (
                  <tr>
                    <td>{this._numberToString(item.Id)}</td>
                    <td>{item.Title}</td>
                    <td>{item.FirstName}</td>
                    <td>{item.Email}</td>
                    <td>{item.Company}</td>
                    <td>${this._numberToString(item.Salary)}</td>
                  </tr>
                ))}
                </tbody>
              </table>

              { this.state.debug && (<pre>{JSON.stringify(this.state, null, 2)}</pre>) }

              </div>
          </div>
        </div>
      </div>
      </Fabric>
    );
  }
}
