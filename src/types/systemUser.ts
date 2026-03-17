export type SystemUser = {
  systemuserid: string;
  fullname: string;
  domainname: string;
  businessunitid?: {
    businessunitid: string;
    name: string;
  };
  isdisabled: boolean;
  applicationid: string | null;
};
