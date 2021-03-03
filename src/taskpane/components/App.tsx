import * as React from "react";
import { Button, ButtonType, DefaultButton, Dialog, DialogFooter } from "office-ui-fabric-react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
// images references in the manifest
import "../../../assets/icon-16.png";
import "../../../assets/icon-32.png";
import "../../../assets/icon-80.png";
/* global Button, console, Excel, Header, HeroList, HeroListItem, Progress */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
  showDialog: boolean;
  count: number;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
      showDialog: false,
      count: 0
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Achieve more with Office integration"
        },
        {
          icon: "Unlock",
          primaryText: "Unlock features and functionality"
        },
        {
          icon: "Design",
          primaryText: "Create and visualize like a pro"
        }
      ]
    });
  }

  click = async () => {
    try {
      await Excel.run(async context => {
        let currentSheetRange = context.workbook.worksheets.getActiveWorksheet().getUsedRangeOrNullObject();

        if (currentSheetRange) {
          let newWorksheet = context.workbook.worksheets.add("newWorksheet");
          newWorksheet.getRange().copyFrom(currentSheetRange, Excel.RangeCopyType.values);
          await context.sync();
          newWorksheet.activate();
        } else {
          this.setState({ showDialog: true });
        }

        await context.sync();
      });
    } catch (error) {
      console.error(error);
    }
  };

  private closeDialog = (): void => {
    this.setState({ showDialog: false });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;
    const { showDialog, count } = this.state;

    if (!isOfficeInitialized) {
      return (
        <Progress title={title} logo="assets/logo-filled.png" message="Please sideload your addin to see app body." />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo="assets/logo-filled.png" title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Office Add-ins can do for you today!" items={this.state.listItems}>
          <p className="ms-font-l">
            Modify the source file, then click <b>Copy</b>.
          </p>
          <Button
            className="ms-welcome__action"
            buttonType={ButtonType.hero}
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.click}
          >
            Copy to a new worksheet
          </Button>
        </HeroList>
        {showDialog && (
          <Dialog hidden={!showDialog} onDismiss={this.closeDialog}>
            <p>Add some data to worksheet to copy!{count}</p>
            <DialogFooter>
              <DefaultButton onClick={this.closeDialog} text="Cancel" />
            </DialogFooter>
          </Dialog>
        )}
      </div>
    );
  }
}
