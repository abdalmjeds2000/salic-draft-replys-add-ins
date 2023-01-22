import * as React from "react";
import PropTypes from "prop-types";
import Header from "./Header";
import DraftReply from "./DraftReply";
import Progress from "./Progress";
import { Pivot, PivotItem, PrimaryButton } from "@fluentui/react";
import accessToken from "../api/accessToken";
import MyDraftReplys from "./MyDraftReplys";

/* global require */

export default class App extends React.Component {
  constructor(props, context) {
    super(props, context);
    this.state = {};
  }

  click = async () => {
    const token = await accessToken();
    if (token) {
      console.log("SUCCESS FETCH TOKEN");
    } else {
      console.log("FAILED FETCH TOKEN");
    }
    console.log('token => ', token);
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return <Progress title={title} logo={require("./../../../assets/salic-logo.png")} message="Please Wait..." />;
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/salic-logo.png")} title={this.props.title} />
        <Pivot aria-label="Count and Icon Pivot Example">
          <PivotItem headerText="Draft Reply" itemIcon="ArrangeByFrom">
            <DraftReply message="Discover what Office Add-ins can do for you today!" items={this.state.listItems} />
          </PivotItem>
          <PivotItem headerText="My Draft Replys" itemIcon="Archive">
            <MyDraftReplys />
          </PivotItem>
        </Pivot>
        <br />
        <br />
        <br />
        <hr />
        <PrimaryButton text="fetch token" onClick={this.click} />
      </div>
    );
  }
}

App.propTypes = {
  title: PropTypes.string,
  isOfficeInitialized: PropTypes.bool,
};
