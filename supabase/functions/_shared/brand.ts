// App and company names, from brand.json at the repo root (also used by the web pages).
import brand from "../../../brand.json" with { type: "json" };

export const APP_NAME: string = brand.appName;
export const APP_TAGLINE: string = brand.tagline;
export const COMPANY_NAME: string = brand.companyName;
