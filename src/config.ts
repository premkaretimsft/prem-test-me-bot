// Force the bot to send all responses through the Canary Bot Framework instance
// regardless of the serviceUrl on the incoming activity.
export const CANARY_SERVICE_URL = "https://canary.botapi.skype.com/amer/";

const config = {
  botId: process.env.BOT_ID,
  botPassword: process.env.BOT_PASSWORD,
  botTenantId: process.env.BOT_TENANT_ID,
  botEndPoint: process.env.BOT_ENDPOINT,
  storageAccountConnectionString: process.env.STORAGE_ACCOUNT_CONNECTION_STRING
};

export default config;
