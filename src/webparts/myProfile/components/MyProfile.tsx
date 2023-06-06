import * as React from "react";
//import styles from './MyProfile.module.scss';
import { IMyProfileProps } from "./IMyProfileProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import styles from "./MyProfile.module.scss";

interface IMyDetails {
  displayName: string;
  mail: string;
  mobilePhone: string;
  id: string;
  type:any;
}

//All ITems interface
/* interface IAllItems {
  AllDetails: IMyDetails[];
} */

export default class MyMail extends React.Component<
  IMyProfileProps,
  IMyDetails
> {
  constructor(props: IMyProfileProps, state: IMyDetails) {
    super(props);
    this.state = { displayName: "", mail: "", mobilePhone: "", id: "" ,type:""};
  }
  componentDidMount(): void {
    this.getMyProfile();
    this.getMyProfilePhoto();
  }

  public getMyProfilePhoto = () => {
    // below code is default code to get
    this.props.context.msGraphClientFactory //default syntax
      .getClient("3") //version updated to 3
      .then((client: MSGraphClientV3): void => {
        client
          .api("me/photo/$value") 
          .version("v1.0")
          .select(" type")
          .get((err: any, res: any) => {
            this.setState({
            /*  type:res.type */
            });
             /* console.log(res);
            console.log(err);  */ 
          });
      });
  };
  public getMyProfile = () => {
    // below code is default code to get
    this.props.context.msGraphClientFactory //default syntax
      .getClient("3") //version updated to 3
      .then((client: MSGraphClientV3): void => {
        client
          .api("me") //to get messages
          .version("v1.0")
          .select(" displayName,mail,mobilePhone,id")
          .get((err: any, res: any) => {
            this.setState({
              displayName: res.displayName,
              mail: res.mail,
              mobilePhone: res.mobilePhone,
              id: res.id,
            });
             console.log(res);
            console.log(err);  
          });
      });
  };

  public render(): React.ReactElement<IMyProfileProps> {
    return (
      <div className={styles.main}>
        <p><h3> My Profile</h3></p>
        <div className={styles.Info}>
        <p><b> {this.state.displayName}</b></p>
        <p>ID: {this.state.id}</p>
        <p>{this.state.mail}</p>
        <p >{this.state.mobilePhone}</p>
        </div>
      </div>
    );
  }
}
