import * as React from "react";
import { PrimaryButton } from "office-ui-fabric-react";
import Script from "react-load-script";

export default class GooglePlacesApi extends React.Component {
  constructor(props) {
    super(props);
  }

  getPlaceIdFromQuery = query => {
    return new Promise(function(resolve, reject) {
      const request = { query: query, fields: ["place_id"] };
      const map = new google.maps.Map(document.getElementById("map"));
      const service = new google.maps.places.PlacesService(map);
      service.findPlaceFromQuery(request, function(results, status) {
        if (status === google.maps.places.PlacesServiceStatus.OK) {
          resolve(results[0].place_id);
        }
        reject(status);
      });
    });
  };

  getDetailsFromPlaceId = async placeId => {
    return new Promise(function(resolve, reject) {
      const request = {
        placeId: placeId,
        fields: ["name", "formatted_address", "website", "formatted_phone_number", "geometry"]
      };
      const map = new google.maps.Map(document.getElementById("map"));
      const service = new google.maps.places.PlacesService(map);
      service.getDetails(request, function(details, status) {
        if (status === google.maps.places.PlacesServiceStatus.OK) {
          resolve(details);
        }
        reject(status);
      });
    });
  };

  search = async () => {
    // console.log('ACTION: "Search" was clicked.');
    let counter = 0;
    try {
      const numb_rows_selected = Number(document.getElementById("rows_selected").innerText);
      if (numb_rows_selected > 1) {
        const selected_row_index = Number(document.getElementById("selected_row_index").innerText);
        for (let i = selected_row_index; i < selected_row_index + numb_rows_selected; i++) {
          // Skip headers.
          if (i === 0) {
            continue;
          }

          const queryValues = await this.buildQueryFromRowIndex(i);
          const query = queryValues.join(" ");

          // Update UI. 
          document.getElementById("query_input").innerText = query;
          counter += 1;
          document.getElementById("searching").innerText =
            "Searching for row " + (i + 1) + ", " + counter + "/" + numb_rows_selected + ".";

          if (query.length > 0) {
            await this.getDetailsFromQuery(i, query);
          } else {
            document.getElementById("searching").innerText = "";
            console.error("Query was empty");
            return;
          }
        }
        document.getElementById("searching").innerText = "";
        return;
      }
    } catch (error) {
      console.error(error);
    }

    const query = document.getElementById("query_input").innerText;
    if (query.length > 0) {
      let selectedRowIndex = Number(document.getElementById("selected_row_index").innerText);
      if (selectedRowIndex === 0) {
        document.getElementById("searching").innerText = "";
        return;
      }
      await this.getDetailsFromQuery(selectedRowIndex, query);
    } else {
      document.getElementById("searching").innerText = "";
      console.error("Query was empty");
      return;
    }
    document.getElementById("searching").innerText = "";
  };

  getQueryColumnAddresses = () => {
    const table = document.getElementById("query_columns_table");
    let addresses = [];
    for (let i = 0; i < table.rows.length; i++) {
      addresses.push(table.rows[i].cells[0].innerText);
    }
    return addresses;
  };

  buildQueryFromRowIndex = async rowIndex => {
    const addresses = this.getQueryColumnAddresses();
    let queryValues = [];

    try {
      await Excel.run(async context => {
        for (const address of addresses) {
          const val = await Excel.run(async context => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const column = sheet.getRange(address).load("columnIndex");
            await context.sync();
            const cell = sheet.getRangeByIndexes(rowIndex, column.columnIndex, 1, 1).load("values");
            await context.sync();
            return cell.values[0][0];
          });
          queryValues.push(val);
        }
      });
    } catch (error) {
      console.error(error);
    }

    return queryValues;
  };

  getDetailsFromQuery = async (row, query) => {
    let details = {};
    let placeId = null;
    let errorMsg = null;
    try {
      placeId = await this.getPlaceIdFromQuery(query);
    } catch (error) {
      console.error("GooglePlacesApi::getPlaceIdFromQuery - " + error);
      errorMsg = error;
    }

    if (!errorMsg) {
      try {
        details = await this.getDetailsFromPlaceId(placeId);
      } catch (error) {
        console.error("GooglePlacesApi::getDetailsFromPlaceId - " + error);
        errorMsg = error;
      }
    }

    this.writeDetailsToTable(row, details, errorMsg);
  };

  writeDetailsToTable = async (row, details, errorMsg) => {
    if (errorMsg) {
      try {
        await Excel.run(async context => {
          const binding = context.workbook.bindings.getItem("name_col");
          const bindingRange = binding.getRange();
          bindingRange.load(["address", "columnIndex"]);
          await context.sync();

          const cell = bindingRange.getCell(row, 0);
          cell.load(["address", "values"]);
          await context.sync();

          cell.values = errorMsg;
          return context.sync();
        });
      } catch (error) {
        console.error("GooglePlacesApi::writeDetailsToTable error...");
        console.error(error);
      }
      return;
    }

    const headers = {
      name: "name",
      address: "formatted_address",
      phone: "formatted_phone_number",
      website: "website",
      latitude: "lat",
      longitude: "lng"
    };

    try {
      Object.entries(headers).map(async ([key, value]) => {
        await Excel.run(async context => {
          const binding = context.workbook.bindings.getItem(key + "_col");
          const bindingRange = binding.getRange();
          bindingRange.load(["address", "columnIndex"]);
          await context.sync();

          const cell = bindingRange.getCell(row, 0);
          cell.load(["address", "values"]);
          await context.sync();

          let newValue = details[value];
          if (value === "lat") {
            newValue = details.geometry.location.lat();
          }
          if (value === "lng") {
            newValue = details.geometry.location.lng();
          }
          cell.values = newValue;

          cell.format.autofitColumns();
          return context.sync();
        });
      });
    } catch (error) {
      console.error("GooglePlacesApi::writeDetailsToTable error...");
      console.error(error);
    }
  };

  render() {
    return (
      <div className="section">
        <div id="map" />
        <PrimaryButton onClick={this.search} iconProps={{ iconName: "ChevronRight" }}>
          Search
        </PrimaryButton>
        <div id="searching" />
        <Script url={"https://maps.googleapis.com/maps/api/js?libraries=places&key=" + this.props.apiKey} />
      </div>
    );
  }
}
