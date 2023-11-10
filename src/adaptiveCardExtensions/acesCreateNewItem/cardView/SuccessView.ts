import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IActionArguments,
} from "@microsoft/sp-adaptive-card-extension-base";
import {
    CARD_VIEW_REGISTRY_ID,
  IAcesCreateNewItemAdaptiveCardExtensionProps,
  IAcesCreateNewItemAdaptiveCardExtensionState,
} from "../AcesCreateNewItemAdaptiveCardExtension";

export class SuccessCardView extends BaseComponentsCardView<
  IAcesCreateNewItemAdaptiveCardExtensionProps,
  IAcesCreateNewItemAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    return BasicCardView({
      cardBar: {
        componentName: "cardBar",
        title: this.properties.title,
      },
      header: {
        componentName: "text",
        text: this.properties.successTxt,
      },
      footer: {
        componentName: "cardButton",
        title: "Back",
        action: {
          type: "Submit",
          parameters: {
            id: "success",
          },
        },
      },
    });
  }

  public async onAction(action: IActionArguments): Promise<void> {
      if (action.type === "Submit" && action.data?.id === "success") {
        this.cardNavigator.replace(CARD_VIEW_REGISTRY_ID)
      }
  }
}
