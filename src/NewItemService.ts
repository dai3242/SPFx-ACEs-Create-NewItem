import { MSGraphClientV3 } from "@microsoft/sp-http";
import { AdaptiveCardExtensionContext } from "@microsoft/sp-adaptive-card-extension-base";

export class NewItemService {
  private MSGraphClient: MSGraphClientV3;
  public context: AdaptiveCardExtensionContext;

  public async _getClient(
    context: AdaptiveCardExtensionContext
  ): Promise<MSGraphClientV3> {
    if (this.MSGraphClient === undefined) {
      this.MSGraphClient = await context.msGraphClientFactory.getClient("3");
    }
    return this.MSGraphClient;
  }

  public setup(context: AdaptiveCardExtensionContext): void {
    this.context = context;
  }

  public async _createItem(
    listId: string,
    itemTitle: string,
    siteId: string | undefined = undefined
  ): Promise<void> {
    const listItem = {
      fields: {
        Title: itemTitle,
      },
    };
    if (siteId) {
      await this.MSGraphClient.api(
        `/sites/${siteId}/lists/${listId}/items`
      ).post(listItem);
    } else {
      await this.MSGraphClient.api(
        `/sites/${siteId}/lists/${listId}/items`
      ).post(listItem);
    }
  }

  public async _getListId(siteId: string, listTitle: string): Promise<string> {
    const list = await this.MSGraphClient.api(
      `/sites/${siteId}/lists/${listTitle}`
    ).get();
    return list.id;
  }

  public async _getSiteId(siteTitle: string): Promise<string> {
    const site = await this.MSGraphClient.api(
      `/sites?search=${siteTitle}`
    ).get();

    return site.value[0].id;
  }
}

const ItemService: NewItemService = new NewItemService();
export default ItemService;
