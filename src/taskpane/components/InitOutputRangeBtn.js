import * as React from "react";
import { DefaultButton } from "office-ui-fabric-react";

export default class InitOutputRangeBtn extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      hidden: false
    };
  }

  initRange = async () => {
    console.log("ACTION: 'Initialize Table' Button was clicked.");
    const headers = new Map([
      ["longitude", "Longitude"],
      ["latitude", "Latitude"],
      ["website", "Website"],
      ["phone", "Phone"],
      ["address", "Address"],
      ["name", "Name"]
    ]);

    try {
      for (const [machineName, displayName] of headers) {
        await Excel.run(async context => {
          // Protect against duplicating bindings.
          let binding;
          try {
            binding = context.workbook.bindings.getItem(machineName + "_col");
            await context.sync();
          } catch (error) {
            console.error(error);
            if (error.code === "InvalidBinding") {
              Office.context.document.bindings.releaseByIdAsync(machineName + "_col");
              await context.sync();
              return;
            }
          }
          try {
            binding.getRange().load(["top", "columnIndex"]);
            await context.sync();
            const hasHeader = sheet.getRangeByIndexes(binding.top, binding.columnIndex, 1, 1).load("values");
            await context.sync();
            if (hasHeader.values.length) {
              console.error(hasHeader.values);
              return;
            }
          } catch (error) {
            if (error.code !== "ItemNotFound") {
              console.error(error);
              return;
            }
          }

          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getUsedRange().load(["rowIndex", "columnIndex"]);
          await context.sync();
          const firstColumn = sheet
            .getRangeByIndexes(range.rowIndex, range.columnIndex, 1, 1)
            .getEntireColumn()
            .load("address");
          await context.sync();

          const newCol = firstColumn.insert("right");
          newCol.load(["address", "top", "columnIndex"]);
          await context.sync();

          const bindings = Office.context.document.bindings;
          bindings.addFromNamedItemAsync(newCol.address, "matrix", { id: machineName + "_col" });

          const newHeader = sheet.getRangeByIndexes(newCol.top, newCol.columnIndex, 1, 1).load("values");
          await context.sync();
          newHeader.values = displayName;
          newHeader.format.autofitColumns();
          return context.sync();
        });
      }
    } catch (error) {
      console.error("InitOutputRangeBtn::initRange() error...");
      console.error(error);
      return;
    }

    this.setState((prevState, props) => {
      return { hidden: true };
    });
  };

  render() {
    if (this.state.hidden) {
      return null;
    }
    return (
      <div className="section">
        <div className="instructions">
          <span className="bullet">Step 1.</span>
          Create the columns where found address details will be stored.
        </div>
        <DefaultButton id="init_output_range_btn" onClick={this.initRange} iconProps={{ iconName: "ChevronRight" }}>
          Create output columns 
        </DefaultButton>
      </div>
    );
  }
}
