import { SPFI, spfi, SPFx } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import "@pnp/sp/webs";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-groups/types";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-users";
import "@pnp/sp/site-users/web";
import "@pnp/sp/content-types";
import "@pnp/sp/sputilities";
import "@pnp/sp/sputilities/types";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/profiles";
import "@pnp/sp/search";
import "@pnp/sp/search/query";
import "@pnp/sp/search/suggest";
import "@pnp/sp/search/types";
import "@pnp/sp/presets/all";
import { ISiteGroupInfo, ISiteUser, ISiteUserInfo } from "@pnp/sp/presets/all";

export class PNP {
    public context: WebPartContext;
    public siteRelativeUrl: string;
    public sp:SPFI;

    constructor(context: WebPartContext) {
        this.context = context;
        this.siteRelativeUrl = context.pageContext.web.serverRelativeUrl;
        this.sp = spfi().using(SPFx(this.context));
    }

    public getConfigValue = async (itemTitle:string):Promise<string> => {
        try {
            const configValueQuery = await this.sp.web.lists.getByTitle('GeneralConfig').items.select('*').filter(`Title eq '${itemTitle}'`)();
            const configValue:string = configValueQuery && configValueQuery.length > 0 ? configValueQuery[0].Value : null;
            return configValue;
        } catch (error) {
            console.error(`Error al obtener la configuracion para ${itemTitle}: ${error}`);
            return '';
        }
    }

    public getCurrentUserInfo = async (): Promise<ISiteUserInfo> => {
        const user:ISiteUserInfo = await this.sp.web.currentUser();
        return user;
    }

    public getAllSiteGroups = async ():Promise<ISiteGroupInfo[]> => {
        const siteGroups:ISiteGroupInfo[] = await this.sp.web.siteGroups();
        return siteGroups;
    }

    public getUsersFromGroup = async (groupname:string):Promise<ISiteUserInfo[]> => {
        try {
            const siteUsers:ISiteUserInfo[] = await this.sp.web.siteGroups.getByName(groupname).users();
            return siteUsers;
        } catch (error) {
            console.error("Error al obtener usuarios del grupo", error);
            return [];
        }
    }

    public addUserToGroup = async (groupName:string):Promise<ISiteUser|null> => {
        try {
            const currentUser:ISiteUserInfo = await this.getCurrentUserInfo();
            const userAdd = await this.sp.web.siteGroups.getByName(groupName).users.add(currentUser.LoginName);
            return userAdd;
        } catch (error) {
            console.error(`Error a intentar agregar usuario al grupo ${groupName}: ${error}`);
            return null;
        }
    }

    public getUserBirthDate = async (userLoginName:string):Promise<string> => {
        try {
            const birthdate = await this.sp.profiles.getUserProfilePropertyFor(userLoginName, 'SPS-Birthday');
            return birthdate;
        } catch (error) {
            console.error(`Error al intentar obtener propiedades del usuario. Error: ${error}`);
            return '';
        }
    }

    public setUserBirthDate = async (userLoginName:string, birthdate:string):Promise<boolean> => {
        try {
            await this.sp.profiles.setSingleValueProfileProperty(userLoginName, 'SPS-Birthday', birthdate);
            return true;
        } catch (error) {
            console.error(`Error al intentar obtener propiedades del usuario. Error: ${error}`);
            return false;
        }
    }
/* 
    public filterUsersWithBirthdayThisMonth = (users: any[]):any[] => {
        const currentMonth = new Date().getMonth();
        const usersWithBirthdayThisMonth = users.filter((user) => {
            const birthday = this.getUserBirthdayInfo(user);
            if (birthday) {
            const birthdayDate = new Date(birthday);
            return birthdayDate.getMonth() === currentMonth;
            }
            return false;
        });

        return usersWithBirthdayThisMonth;
    } */
}