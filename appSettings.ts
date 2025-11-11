const settings: AppSettings = {
  'clientId': process.env.clientId,
  'clientSecret': process.env.clientSecret,
  'tenantId': process.env.tenantId
};

export interface AppSettings {
  clientId: string;
  clientSecret: string;
  tenantId: string;
}

export default settings;