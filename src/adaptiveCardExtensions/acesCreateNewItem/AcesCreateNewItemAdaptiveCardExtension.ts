import type { IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseAdaptiveCardExtension } from "@microsoft/sp-adaptive-card-extension-base";
import { CardView } from "./cardView/CardView";
import { AcesCreateNewItemPropertyPane } from "./AcesCreateNewItemPropertyPane";
// import { SuccessCardView } from "./cardView/SuccessView";
// import { ErrorCardView } from "./cardView/ErrorView";
import ItemService from "../../NewItemService";
import { SuccessCardView } from "./cardView/SuccessView";
import { ErrorCardView } from "./cardView/ErrorView";

export interface IAcesCreateNewItemAdaptiveCardExtensionProps {
  title: string;
  listTitle: string;
  siteTitle: string;
  buttonLabel: string;
  subTitle: string;
  iconName: string;
  successTxt: string;
  errorTxt: string;
}

export interface IAcesCreateNewItemAdaptiveCardExtensionState {}

export const CARD_VIEW_REGISTRY_ID: string = "AcesCreateNewItem_CARD_VIEW";
export const SUCCESS_CARD_VIEW_REGISTRY_ID: string =
  "AcesCreateNewItem_SUCCESS_CARD_VIEW";
export const ERROR_CARD_VIEW_REGISTRY_ID: string =
  "AcesCreateNewItem_ERROR_CARD_VIEW";

export default class AcesCreateNewItemAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IAcesCreateNewItemAdaptiveCardExtensionProps,
  IAcesCreateNewItemAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: AcesCreateNewItemPropertyPane;

  public async onInit(): Promise<void> {
    this.state = {};

    // registers the card view to be shown in a dashboard
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.cardNavigator.register(
      SUCCESS_CARD_VIEW_REGISTRY_ID,
      () => new SuccessCardView()
    );
    this.cardNavigator.register(
      ERROR_CARD_VIEW_REGISTRY_ID,
      () => new ErrorCardView()
    );

    await ItemService._getClient(this.context);
    ItemService.setup(this.context);

    console.log("context", this.context)

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'AcesCreateNewItem-property-pane'*/
      "./AcesCreateNewItemPropertyPane"
    ).then((component) => {
      this._deferredPropertyPane =
        new component.AcesCreateNewItemPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }

  protected async onPropertyPaneFieldChanged(
    propertyPath: string,
    oldValue: any,
    newValue: any
  ): Promise<void> {
    if (newValue !== oldValue) {
      // if (propertyPath === "listTitle") {

      // }
      this.renderCard();
    }
  }
}
