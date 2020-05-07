import * as React from "react";
import { PrimaryButton } from "office-ui-fabric-react";

export default class BuildJsonBtn extends React.Component {

  getAllRowsWithNameValue = async () => {
    try {
      return await Excel.run(async context => {
        const nameCol = context.workbook.bindings.getItem("name_col");
        const nameRange = nameCol.getRange().getUsedRange();
        nameRange.load(["address", "top", "rowIndex", "rowCount", "values"]);
        await context.sync();

        let rows = [];
        for (let i = nameRange.top + 1; i < nameRange.rowCount; i++) {
          const value = nameRange.values[i][0];
          if (value.length && value !== "ZERO_RESULTS" && value.length && value !== "REQUEST_DENIED") {
            rows.push(i);
          }
        }
        return rows;
      });
    } catch (error) {
      console.error(error);
    }
  };

  buildJson = async () => {
    console.log('ACTION: "Export as JSON" was clicked.');
    document.getElementById("json_output").value = "Building JSON...";
    let places = new Map();
    const rows = await this.getAllRowsWithNameValue();
    const headers = new Map([
      ["name", true],
      ["address", true],
      ["phone", false],
      ["website", false],
      ["latitude", true],
      ["longitude", true]
    ]);

    try {
      for (const row of rows) {
        let place = {};
        place.id = row;

        for (const [machineName, required] of headers) {
          let transport = await Excel.run(async context => {
            const binding = context.workbook.bindings.getItem(machineName + "_col");
            const bindingRange = binding.getRange();
            const cell = bindingRange.getCell(row, 0);
            await context.sync();
            cell.load("values");
            await context.sync();

            const value = cell.values[0][0];

            let transport = {};
            transport[machineName] = value;

            return transport;
          });
          place = { ...place, ...transport };
        }

        places.set(place.name, place);
      }
      console.log(places);
      places = new Map([...places.entries()].sort());
      let sortedPlaces = [];
      for (const [key, value] of places) {
        sortedPlaces.push(value);
      }
      console.log(sortedPlaces);

      document.getElementById("json_output").value = JSON.stringify(sortedPlaces);
    } catch (error) {
      console.error(error);
    }
  };

  render() {
    return (
      <div className="section">
        <div className="instructions">
          <span className="bullet">Step 5.</span>
          Export found Place Details as JSON.
        </div>
        <PrimaryButton onClick={this.buildJson} iconProps={{ iconName: "ChevronRight" }}>
          Build JSON
        </PrimaryButton>
        <input id="json_output" type="textarea" readOnly={true} style={{ width: "280px" }} />
      </div>
    );
  }
}
