import * as React from "react";
// import styles from "./Projects.module.scss";
import { IProjectsProps } from "./ProjectsProps";
import { IProjectsState } from "./ProjectState";
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
// import {
//   DocumentCard,
//   DocumentCardDetails,
//   DocumentCardTitle,
// } from "office-ui-fabric-react/lib/DocumentCard";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import {
  DocumentCard,
  DocumentCardActivity,
  DocumentCardPreview,
  DocumentCardTitle,
  DocumentCardDetails,
  IDocumentCardPreviewProps,
} from "@fluentui/react/lib/DocumentCard";
import { ImageFit } from "@fluentui/react/lib/Image";

// export interface IProjectsProps {}

// export interface IProjectsState {}

export default class Projects extends React.Component<
  IProjectsProps,
  IProjectsState
> {
  constructor(props: IProjectsProps, state: IProjectsState) {
    super(props);

    this.state = {
      items: [],
    };
  }

  public getItems(filterVal) {
    //   this.context.httpClient.get("https://your-web-api", HttpClient.configurations.v1)
    //   .then((data: HttpClientResponse) => data.json())
    //   .then((data: any) => {

    //   });
    if (filterVal === "*") {
      this.props.context.spHttpClient
        .get(
          `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/Items?$expand=ProjectManager&$select=*,ProjectManager,ProjectManager/EMail,ProjectManager/Title`,
          SPHttpClient.configurations.v1
        )
        .then(
          (response: SPHttpClientResponse): Promise<{ value: any[] }> => {
            return response.json();
          }
        )
        .then((response: { value: any[] }) => {
          var _items = [];
          _items = _items.concat(response.value);
          this.setState({
            items: _items,
          });
        });
    } else {
      this.props.context.spHttpClient
        .get(
          `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Projects')/Items?$expand=ProjectManager&$select=*,ProjectManager,ProjectManager/EMail,ProjectManager/Title&$filter=Status eq %27${filterVal}%27`,
          SPHttpClient.configurations.v1
        )
        .then(
          (response: SPHttpClientResponse): Promise<{ value: any[] }> => {
            return response.json();
          }
        )
        .then((response: { value: any[] }) => {
          var _items = [];
          _items = _items.concat(response.value);
          this.setState({
            items: _items,
          });
        });
    }
  }
  public componentDidMount() {
    var getAll = "*";
    this.getItems(getAll);
  }
  public progFilter(filterVal) {
    switch (filterVal) {
      case "All":
        return this.getItems(filterVal);
      case "In Progress":
        return this.getItems(filterVal);
      case "Not Started":
        return this.getItems(filterVal);
      case "Completed":
        return this.getItems(filterVal);
      case "On Hold":
        return this.getItems(filterVal);
      default:
        return this.getItems(filterVal);
    }
  }

  public render(): React.ReactElement<IProjectsProps> {
    var _projDocLink = `${this.props.context.pageContext.web.absoluteUrl}/Project%20Documents/Forms/AllItems.aspx?FilterField1=Project&FilterValue1=`;
    var notStarted = "Not Started";
    var inProg = "In Progress";
    var comp = "Completed";
    var onHold = "On Hold";
    var allPrj = "*";

    return (
      <div>
        <div>
          <PrimaryButton
            text="All"
            onClick={() => this.progFilter(allPrj)}
            allowDisabledFocus
          />
          <PrimaryButton
            text="In Progress"
            onClick={() => this.progFilter(inProg)}
            allowDisabledFocus
          />
          <PrimaryButton
            text="Completed"
            onClick={() => this.progFilter(comp)}
            allowDisabledFocus
          />
          <PrimaryButton
            text="On Hold"
            onClick={() => this.progFilter(onHold)}
            allowDisabledFocus
          />
          <PrimaryButton
            text=" Not Started"
            onClick={() => this.progFilter(notStarted)}
            allowDisabledFocus
          />
        </div>
        {this.state.items.map((item, key) => (
          <DocumentCard>
            {/* <a href={_projDocLink + item.Title} target="_blank">
              <DocumentCardTitle title={item.Title}></DocumentCardTitle>
            </a> */}
            {/* <DocumentCardTitle title={item.Title}></DocumentCardTitle> */}
            <a href={_projDocLink + item.Title} target="_blank">
              <DocumentCardTitle title={item.Title}></DocumentCardTitle>
            </a>

            <a href={"mailto:" + item.ProjectManager.EMail}>
              <DocumentCardTitle
                title={item.ProjectManager.Title}
                showAsSecondaryTitle
              ></DocumentCardTitle>
            </a>

            <DocumentCardTitle
              title={item.Status}
              showAsSecondaryTitle
            ></DocumentCardTitle>
          </DocumentCard>
        ))}
      </div>
    );
  }
}
