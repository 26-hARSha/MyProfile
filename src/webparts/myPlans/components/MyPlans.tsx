import * as React from "react";
//import styles from './MyPlans.module.scss';
import { IMyPlansProps } from "./IMyPlansProps";
//import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as moment from "moment";
import styles from "./MyPlans.module.scss";

interface IPlans {
  title: string;
  dueDateTime: string;
  assignedBy: {
    user: {
        displayName: string;
        id: string;
    },
    application: {
      displayName: string;
      id: string;
  };
}
hasDescription:string;
priority:string;
percentComplete:string;
}

//All ITems interface
interface IAllItems {
  AllPlans: IPlans[];
}

export default class MyMail extends React.Component<IMyPlansProps, IAllItems> {
  constructor(props: IMyPlansProps, state: IAllItems) {
    super(props);
    this.state = {
      AllPlans: [],
    };
  }
  componentDidMount(): void {
    this.getMyPlans();
  }

  public getMyPlans = () => {
 
 /*  let top = this.props.noofPlans; */

    // below code is default code to get
    this.props.context.msGraphClientFactory                                         //default syntax
      .getClient("3") 
      .then((client: MSGraphClientV3): void => {
        client
          .api("me/planner/tasks") //to get messages
          .version("v1.0")
          /* .top("") */
          .select("dueDateTime,title,percentComplete,hasDescription,priority")            // selected columns from response preview
          .get((err: any, res: any) => {  
            this.setState({
              AllPlans: res.value,
            });
            // console.log(this.state.AllPlans);                                                     //checking
           /*  console.log(res);
            console.log(err); */
          });
      });
  };

  public render(): React.ReactElement<IMyPlansProps> {
    return (
      <div className={styles.main}
      style={{
        height: this.props.webHeight,
      }}>
        <h3>My Plans</h3>
        {this.state.AllPlans.map((plan) => {
          return (
            <><div className={styles.planDetails}>

              <p><b>{plan.title}</b></p>
              <p>{moment(plan.dueDateTime).format("LL")}</p>
              <p>Task Complete : {plan.percentComplete} %</p>
              {/* <p>{plan.hasDescription}</p>
              <p>{plan.priority}</p> */}
              </div>
            </> 
          );
        })}
      </div>
    );
  }
}
