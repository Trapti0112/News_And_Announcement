export default class spcommon {
    public static fIsNullOrUndefined(obj: any, checkData: boolean): boolean {
        if (obj == undefined || obj == null)
            return true;
        else if (checkData)
            if (obj.length <= 0)
                return true;
            else
                return false;
        else
            return false;
    }


    public static ChangeDateFormat(publishDate: string): React.ReactNode {
        if (publishDate != "") {
            let dt = new Date(publishDate);
            return dt.getDate() + "/" + (dt.getMonth() + 1) + "/" + dt.getFullYear();
        } else
            return "";
    }
}


