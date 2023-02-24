import { sp, Web } from '@pnp/sp';
import { WebPartContext } from "@microsoft/sp-webpart-base";
// npm install @pnp/logging@2.0.10 @pnp/common @1.3.4 @pnp/odata@1.3.2 --save
// @pnp/sp@1.3.2 --save
export default class spservices {
    constructor(private context: WebPartContext) {
        // Setuo Context to PnPjs and MSGraph
        sp.setup({
            spfxContext: this.context
        });
        //this.onInit();
    }


    public async getListItems(siteUrl: string, listTitle: string, selectQuery: string, expandQuery: string, filterQuery: string, numberImages: number, orderBy: string, isAscending: boolean): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);

            results = await web.lists
                .getByTitle(listTitle).items
                .select(selectQuery)
                .top(numberImages)
                .expand(expandQuery)
                .filter(filterQuery)
                .orderBy(orderBy, isAscending)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getListItemsByListId(siteUrl: string, listId: string, selectQuery: string, expandQuery: string, filterQuery: string, numberImages: number, orderBy: string, isAscending: boolean): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getById(listId).items
                .select(selectQuery)
                .top(numberImages)
                .expand(expandQuery)
                .filter(filterQuery)
                .orderBy(orderBy, isAscending)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getListItemByID(siteUrl: string, listTitle: string, selectQuery: string, itemID: number): Promise<any> {
        let results: any = {};
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getByTitle(listTitle).items.getById(itemID)
                .select(selectQuery)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getListItemProperties(siteUrl: string, listTitle: string, selectQuery: string, itemID: number): Promise<any> {
        let results: any = {};
        try {
            const web = new Web(siteUrl);

            results = await web.lists
                .getByTitle(listTitle)
                .items.getById(itemID)

                .select(selectQuery)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async updateListItem(siteUrl: string, listTitle: string, itemID: number, itemData: any): Promise<any> {
        let results: any = {};
        try {
            const web = new Web(siteUrl);
            //$pnp.sp.web.lists.getById(listId).items.getById(ItemData.Id).update(ItemData)
            results = await web.lists
                .getByTitle(listTitle)
                .items
                .getById(itemID)
                .update(itemData);
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async addListItem(siteUrl: string, listTitle: string, itemData: any): Promise<any> {
        let results: any = {};
        try {
            const web = new Web(siteUrl);
            //$pnp.sp.web.lists.getById(listId).items.getById(ItemData.Id).update(ItemData)
            results = await web.lists
                .getByTitle(listTitle)
                .items
                .add(itemData);
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async deleteListItem(siteUrl: string, listTitle: string, itemId: number): Promise<any> {
        let results: any = {};
        try {
            const web = new Web(siteUrl);
            //$pnp.sp.web.lists.getById(listId).items.getById(ItemData.Id).update(ItemData)
            results = await web.lists
                .getByTitle(listTitle)
                .items
                .getById(itemId)
                .delete();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getImages(siteUrl: string, listTitle: string, numberImages: number): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getByTitle(listTitle).items
                .select('Title', 'Description', 'File_x0020_Type', 'FileSystemObjectType', 'File/Name', 'File/ServerRelativeUrl', 'File/Title', 'File/Id', 'File/TimeLastModified')
                .top(numberImages)
                .expand('File')
                .filter((`File_x0020_Type eq  'jpg' or File_x0020_Type eq  'png' or  File_x0020_Type eq  'jpeg'  or  File_x0020_Type eq  'gif' or  File_x0020_Type eq  'mp4'`))
                .orderBy('Id')
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getListFieldByTitle(siteUrl: string, listTitle: string, fieldTitle: string): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getByTitle(listTitle).fields.getByInternalNameOrTitle(fieldTitle)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getListFields(siteUrl: string, listId: string, filterQuery: string): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getByTitle(listId).fields
                .filter(filterQuery)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getListPropertiesByListTitle(siteUrl: string, listTitle: string, selectQuery: string, expandQuery: string): Promise<any> {
        let results: any = {};
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getByTitle(listTitle)
                .select(selectQuery)
                .expand(expandQuery)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        // sort by name

        return results;
    }

    public async getWebDetails(siteUrl: string, selectQuery: string): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web
                .select(selectQuery)
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        return results;
    }

    public async getLibraryFolders(siteUrl: string, listTitle: string): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getByTitle(listTitle)
                .rootFolder
                .folders
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        return results;
    }

    public async getLibraryFolderSubFolders(siteUrl: string, listTitle: string, folderName: string): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web.lists.getByTitle(listTitle).rootFolder.folders.getByName(folderName).folders.get();
        } catch (error) {
            return Promise.reject(error);
        }
        return results;
    }

    public async getLibraryDefaultView(siteUrl: string, listTitle: string): Promise<any[]> {
        let results: any[] = [];
        try {
            const web = new Web(siteUrl);
            results = await web.lists
                .getByTitle(listTitle)
                .defaultView
                .get();
        } catch (error) {
            return Promise.reject(error);
        }
        return results;
    }

}