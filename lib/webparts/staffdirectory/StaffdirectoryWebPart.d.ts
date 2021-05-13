import { Version } from "@microsoft/sp-core-library";
import { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import "@pnp/graph/users";
import "@pnp/graph/photos";
import "@pnp/sp/profiles";
import "../../ExternalRef/CSS/style.css";
export interface IStaffdirectoryWebPartProps {
    description: string;
}
export default class StaffdirectoryWebPart extends BaseClientSideWebPart<IStaffdirectoryWebPartProps> {
    onInit(): Promise<void>;
    render(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=StaffdirectoryWebPart.d.ts.map