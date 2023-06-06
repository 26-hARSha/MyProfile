import * as React from 'react';
//import styles from './Events.module.scss';
import { IEventsProps } from './IEventsProps';
import { MSGraphClientV3 } from "@microsoft/sp-http";
//import * as moment from "moment";


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

export default class MyMail extends React.Component<IEventsProps, IAllItems> {
  constructor(props: IEventsProps, state: IAllItems) {
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
             console.log(res);
            console.log(err); 
          });
      });
  };

  public render(): React.ReactElement<IEventsProps> {
    return (
     <div>My Events</div>
    );
  }
}
