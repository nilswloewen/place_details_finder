import * as React from "react";
import Script from "react-load-script";

export default class ApiKeyForm extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      apiKey: "Paste Key Here"
    };

    this.handleChange = this.handleChange.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
  }

  handleChange(event) {
    this.setState({ apiKey: event.target.value });
  }

  getKey = async () => {
    const key = await OfficeRuntime.storage.getItem("apiKey").then(result => {
      return result;
    });

    if (typeof key !== "undefined" && key.length() > 0) {
      return key;
    }
    return null;
  };

  storeKey = async value => {
    const key = "apiKey";
    console.log("storeKey:" + value);
    const report = await OfficeRuntime.storage.setItem(key, value).then(
      function(result) {
        return "Success: Item with key '" + key + "' saved to storage.";
      },
      function(error) {
        return "Error: Unable to save item with key '" + key + "' to storage. " + error;
      }
    );
    console.log(report);
  };

  handleSubmit = async event => {
    event.preventDefault();
    await this.storeKey(this.state.apiKey);
  };

  render() {
    const key = this.getKey(); 
    console.warn(key);
    return (
      <div className="section">
        <div className="instructions">
          <span className="bullet">Step 4.</span>
          Enter your <a href="https://cloud.google.com/maps-platform/">Google Places Api</a> key.
        </div>
        <label>
          ApiKey:
          <input defaultValue={this.state.apiKey} onChange={this.handleChange} />
        </label>
        <input type="submit" value="Submit" onClick={this.handleSubmit} />

        <Script url={"https://maps.googleapis.com/maps/api/js?key=" + this.state.apiKey + "&libraries=places"} />
      </div>
    );
  }
}
 