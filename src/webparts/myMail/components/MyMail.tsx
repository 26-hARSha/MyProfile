import * as React from "react";
import styles from "./MyMail.module.scss";
import { IMyMailProps } from "./IMyMailProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as moment from "moment";

//Email Intefaces
interface IEmails {
  emailAddres: any;
  subject: string;
  webLink: string; // mail Link
  from: {
    emailAddress: {
      name: string;
      address: string;
    };
  };
  receivedDateTime: any;
  bodyPreview: string;
  isRead: any;
}

//All ITems interface
interface IAllItems {
  AllEmails: IEmails[];
}

export default class MyMail extends React.Component<IMyMailProps, IAllItems> {
  constructor(props: IMyMailProps, state: IAllItems) {
    super(props);
    this.state = {
      AllEmails: [],
    };
  }
  componentDidMount(): void {
    this.getMyEmails();
  }

  public getMyEmails = () => {
    /* console.log("test emails");
  alert("My Mails") */

  
    // below code is default code to get
    this.props.context.msGraphClientFactory //default syntax
      .getClient("3") //version updated to 3
      .then((client: MSGraphClientV3): void => {
        client
          .api("me/messages") //to get messages
          .version("v1.0")
          .select("subject,webLink, from,receivedDateTime,isRead,bodyPreview") // selected columns from response preview
          .get((err: any, res: any) => {
            this.setState({
              AllEmails: res.value,
            });
            // console.log(this.state.AllEmails);         //checking
            /*  console.log(res);
            console.log(err); */
          });
      });
  };

  public render(): React.ReactElement<IMyMailProps> {
    return (
      <div className={styles.main}
      style={{
        height: this.props.webHeight,
      }}>
        <h3>My Mails</h3>
        {this.state.AllEmails.map((email) => {
          return (
            <div >
              <div
                className={styles.cardNumber}
               
                style={{
                  backgroundColor: email.isRead == false ? "  rgb(226, 226, 226)" : " rgb(255, 255, 255)",
                 
                   /*  height: this.props.webHeight, */
                  
                }}
              >
         <p className={styles.circleshape} style={{
                  backgroundColor: email.isRead == false ? " red" : " green",
                }}></p> 
                <p> {email.from.emailAddress.name}</p>
                <p>{email.subject}</p>
                <p>{moment(email.receivedDateTime).format("LL")}</p>
                <p>{email.bodyPreview}</p>
                <button
                  onClick={() => {
                    window.open.apply(email.webLink, "_blank");
                  }}
                >
                  {" "}
                  Open Email in new tab
                </button>
                <hr />
              </div>
            </div>
          );
        })}
      </div>
    );
  }
}
