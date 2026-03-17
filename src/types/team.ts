export type Team = {
  teamid: string;
  name: string;
  teamtype: number;
  businessunitid?: {
    businessunitid: string;
    name: string;
  };
  isdefault: boolean;
};
