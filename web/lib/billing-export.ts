export type BillingExport={legalName:string;country:string;foreignTaxId?:string;billingAddress:string;billingEmail:string;purchaseReference:string;amount:string;currency:string;date:string};
export interface BillingExportService{export(record:BillingExport):Promise<{accepted:boolean}>}
export class DisabledBillingExportService implements BillingExportService{async export(){return {accepted:false}}}
