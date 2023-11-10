import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'AcesCreateNewItemAdaptiveCardExtensionStrings';
import {
  IAcesCreateNewItemAdaptiveCardExtensionProps,
  IAcesCreateNewItemAdaptiveCardExtensionState
} from '../AcesCreateNewItemAdaptiveCardExtension';

export interface IQuickViewData {
  subTitle: string;
  title: string;
}

export class QuickView extends BaseAdaptiveCardQuickView<
  IAcesCreateNewItemAdaptiveCardExtensionProps,
  IAcesCreateNewItemAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      subTitle: strings.SubTitle,
      title: strings.Title
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}
