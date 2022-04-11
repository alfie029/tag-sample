import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";

/* global console, Office, PowerPoint, require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };
  }

  componentDidMount() {
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: " Write tags to Presentation document level",
        },
        {
          icon: "Unlock",
          primaryText: " Write tags for each slides",
        },
        {
          icon: "Design",
          primaryText: " Write tags for each shapes",
        },
      ],
    });
  }

  traceStep(data) {
    /**
     * Insert your PowerPoint code here
     */
    Office.context.document.setSelectedDataAsync(
      `Step ${data}`,
      {
        coercionType: Office.CoercionType.Text,
      },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error(result.error.message);
        }
      }
    );
  }

  writeTagToPresentation = async () => {
    await PowerPoint.run(async (context) => {
      context.presentation.tags.add(
        "ContosoInstance",
        JSON.stringify({
          id: "Contoso",
          category: "document",
          value: `test data for document (${Date.now().toString()})`,
        })
      );
      await context.sync();
      this.traceStep("writeTagToPresentation");
    });
  };

  writeTagToSlides = async () => {
    await PowerPoint.run(async (context) => {
      // NOTE, below two lines are required to enumerate slides
      context.presentation.slides.load();
      await context.sync();
      context.presentation.slides.items.forEach((slide, idx) => {
        slide.tags.add(
          "Contoso-Slide",
          JSON.stringify({
            id: "ContosoPage",
            category: "slide",
            value: `test data for slide-${idx} (${Date.now().toString()})`,
          })
        );
      });
      await context.sync();
      this.traceStep("writeTagToSlides");
    });
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <Header logo={require("./../../../assets/logo-filled.png")} title={this.props.title} message="Welcome" />
        <HeroList message="Discover what Tag in Presentation!" items={this.state.listItems}>
          <p className="ms-font-l">
            click below <b>buttons</b> step by step.
          </p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.writeTagToPresentation}
          >
            1. Write Tag to Presentation
          </DefaultButton>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.writeTagToSlides}
          >
            2. Write Tag to Slides
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}
