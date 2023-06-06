import * as React from 'react';

import { IManagerProps } from './IManagerProps';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import styles from '../../myMail/components/MyMail.module.scss';


interface IMyManager {
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
IManagerProps,
  IMyManager
> {
  constructor(props: IManagerProps, state: IMyManager) {
    super(props);
    this.state = { displayName: "", mail: "", mobilePhone: "", id: "" ,type:""};
  }
  componentDidMount(): void {
    this.getMyManager();
  }


  public getMyManager = () => {
    // below code is default code to get
    this.props.context.msGraphClientFactory //default syntax
      .getClient("3") //version updated to 3
      .then((client: MSGraphClientV3): void => {
        client
          .api("me/manager") //to get messages
          .version("v1.0")
          .select(" displayName,mail,mobilePhone,id")
          .get((err: any, res: any) => {
            this.setState({
              displayName: res.displayName,
              mail: res.mail,
              mobilePhone: res.mobilePhone,
              id: res.id,
            });
             /* console.log(res);
            console.log(err);  */ 
          });
      });
  };

  public render(): React.ReactElement<IManagerProps> {
    return (
      <div className={styles.main}>
        <p><h3> My Manager</h3></p>
        <div>
        <p><b> {this.state.displayName}</b></p>
        <p >ID: {this.state.id}</p>
        <p>{this.state.mail}</p>
        <p  style={{ display: this.state.mobilePhone  == null? " not available" : "{this.state.mobilePhone}" }}>Mobile Number: not available</p>
        </div>
      </div>
    );
  }
}
