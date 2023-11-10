import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  IActionArguments,
  // IExternalLinkCardAction,
  // IQuickViewCardAction,
  TextInputCardView,
} from "@microsoft/sp-adaptive-card-extension-base";
import * as strings from "AcesCreateNewItemAdaptiveCardExtensionStrings";
import {
  SUCCESS_CARD_VIEW_REGISTRY_ID,
  ERROR_CARD_VIEW_REGISTRY_ID,
  IAcesCreateNewItemAdaptiveCardExtensionProps,
  IAcesCreateNewItemAdaptiveCardExtensionState,
} from "../AcesCreateNewItemAdaptiveCardExtension";
import ItemService from "../../../NewItemService";

export class CardView extends BaseComponentsCardView<
  IAcesCreateNewItemAdaptiveCardExtensionProps,
  IAcesCreateNewItemAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    return TextInputCardView({
      cardBar: {
        componentName: "cardBar",
        title: this.properties.title,
      },
      header: {
        componentName: "text",
        text: this.properties.subTitle,
      },
      body: {
        componentName: "textInput",
        placeholder: strings.Placeholder,
        id: "item",
        iconBefore: {
          url: "Edit",
        },
      },
      footer: {
        componentName: "cardButton",
        title: this.properties.buttonLabel,
        action: {
          type: "Submit",
          parameters: {
            id: "sendItem",
          },
        },
      },
    });
  }

  public async onAction(action: IActionArguments): Promise<void> {
    try {
      if (action.type === "Submit" && action.data?.id === "sendItem") {
        const item: string = action.data.item;
        const siteId = await ItemService._getSiteId(this.properties.siteTitle);
        const listId = await ItemService._getListId(siteId, this.properties.listTitle);
        ItemService._createItem(listId, item, siteId);
        this.cardNavigator.replace(SUCCESS_CARD_VIEW_REGISTRY_ID)
      }
    } catch (error) {
      console.error(error);
      this.cardNavigator.replace(ERROR_CARD_VIEW_REGISTRY_ID)
    }
  }
  
}
