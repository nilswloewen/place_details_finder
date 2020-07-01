import * as React from "react";
import Progress from "./Progress";
import QueryColumnsTable from "./QueryColumnsTable";
import GooglePlacesApi from "./GooglePlacesApi";
import InitOutputRangeBtn from "./InitOutputRangeBtn";
import BuildJsonBtn from "./BuildJsonBtn";

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {
      apiKey: ""
    };
    this.handleChange = this.handleChange.bind(this);
  }

  componentDidMount = async () => {
    if (!this.state.validKey) {
      setTimeout(function() {}, 5000);
      const error = document.getElementById("google_error");
      if (error) {
        if (error.innerText === "NoErrorReported") {
          this.setState({
            validKey: true
          });
        }
      }
    }
  };

  handleChange(event) {
    let input = event.target.value;
    if (input) {
      this.setState({ apiKey: input.trim() });
    }
  }

  onSelectionChange = async args => {
    console.log("onSelectionChange() fired.");
    let get = async () => {
      await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(args.worksheetId);
        const selectedRange = sheet.getRange(args.address).load(["address", "rowCount", "rowIndex"]);
        await context.sync();

        const query = await this.buildQueryFromRowIndex(context, args.worksheetId, selectedRange.rowIndex);

        // Display selected range and amount of rows.
        document.getElementById("selected_address").innerText = args.address + 1;
        document.getElementById("rows_selected").innerText = selectedRange.rowCount;
        document.getElementById("selected_row_index").innerText = selectedRange.rowIndex;
        document.getElementById("query_input").innerText = query.join(" ");
      });
    };
    get.bind(args);
    get();
  };

  getQueryColumnAddresses = () => {
    const table = document.getElementById("query_columns_table");
    let addresses = [];
    for (let i = 0; i < table.rows.length; i++) {
      addresses.push(table.rows[i].cells[0].innerText);
    }
    return addresses;
  };

  buildQueryFromRowIndex = async (context, worksheetId, rowIndex) => {
    const addresses = this.getQueryColumnAddresses();
    let queryValues = [];

    for (const address of addresses) {
      const val = await Excel.run(async context => {
        const sheet = context.workbook.worksheets.getItem(worksheetId);
        const column = sheet.getRange(address).load("columnIndex");
        await context.sync();
        const cell = sheet.getRangeByIndexes(rowIndex, column.columnIndex, 1, 1).load("values");
        await context.sync();
        return cell.values[0][0];
      });
      queryValues.push(val);
    }

    return queryValues;
  };

  attachSelectionEventToTable = async () => {
    Excel.run(async context => {
      let sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onSelectionChanged.add(this.onSelectionChange);

      await context.sync();
      console.log("onSelectionChanged event successfully registered SourceTable.");
    });
  };

  checkForValidApiKey() {
    const err = document.getElementById("google_error");
    if (err.innerText === "GoogleError") {
      this.setState({ apiKey: "Invalid", validKey: false });
    }
  }

  render() {
    const { title, isOfficeInitialized } = this.props;
    if (!isOfficeInitialized) {
      return <Progress title={title} message="Details Finder is loading..." />;
    }

    if (!this.state.apiKey || !this.state.validKey) {
      return (
          <div className="section">
            <div className="instructions">
              <span className="bullet">Link with Google Places API</span>
              <p>
                This add-in will enable you to find place data, such as place name, address, phone number, website, latitude, and longitude based on partial address data already in your spreadsheet.
              </p>
              <p>
                This add-in does not require an account or for you to login directly. However, you  will need a Google Account and valid credit card information in order to get your own Google Places API key.
                Visit <a href="https://cloud.google.com/maps-platform/" target="_blank" rel="noopener noreferrer">Google Maps Platform</a> to create an account.
              </p>
              <p>
                The Places API uses a pay-as-you-go pricing model. Visit <a href="https://developers.google.com/places/web-service/usage-and-billing" target="_blank" rel="noopener noreferrer">Places API Usage and Billing</a> for more information. This add-in does not deal with your payment information in any way.
              </p>
              <p>
                This add-in does not transfer or store your API key. The key is used only in the current session when you are using this add-in.
              </p>
              <p>
                You may need to enter the API key every time you start this add-in. If you have entered an incorrect key, simply close and remove this add-in, and then re-install and open this add-in to enter the correct key.
              </p>
              <p>
                Once you have a Google Places API Key, please enter it here:
              </p>
            </div>
            <label>
              <input defaultValue={this.state.apiKey} onChange={this.handleChange} placeholder="Paste Key Here" />
            </label>
          </div>
      );
    }

    this.attachSelectionEventToTable();

    return (
        <div>
          <div id="selected_address" className="hidden" />
          <div id="selected_row_index" className="hidden" />
          <div id="rows_selected" className="hidden" />
          <label>
            API Key:
            <div id="apiKey">{this.state.apiKey}</div>
          </label>
          <InitOutputRangeBtn />
          <QueryColumnsTable />
          <div className="section">
            <div className="instructions">
              <span className="bullet">Step 3.</span>
              Select a row then review and/or modify the query built from your selection.
              <p>
                If multiple rows are selected, only the first query will be shown here. When "Search" is clicked, your queries will be sent one at a time automatically.
              </p>
            </div>
            <div contentEditable={true} id="query_input" placeholder={"Click on a row..."} style={{ width: "280px" }} />
          </div>
          <GooglePlacesApi apiKey={this.state.apiKey} />
          <BuildJsonBtn />
        </div>
    );
  }
}
