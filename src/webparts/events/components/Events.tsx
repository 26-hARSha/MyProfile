import * as React from "react";
//import styles from './Events.module.scss';
import { IEventsProps } from "./IEventsProps";
import { MSGraphClientV3 } from "@microsoft/sp-http";
import * as moment from "moment";
import styles from "./Events.module.scss";

//ToDo Intefaces

interface IToDo {
  title: string;
  status: string;
  importance: string;
  isReminderOn: string;
  createdDateTime: string;
  hasAttachments: string;
  categories: string;
  completedDateTime: {
    dateTime: string;
    recurrence: {
      range: {
          type: string;
          startDate: string;
          endDate: String;
          recurrenceTimeZone:string;
          numberOfOccurrences: number;
      }
  };
}}

//All ITems interface
interface IAllItems {
  AllToDos: IToDo[];
}

export default class MyToDo extends React.Component<IEventsProps, IAllItems> {
  constructor(props: IEventsProps, state: IAllItems) {
    super(props);
    this.state = {
      AllToDos: [],
    };
  }
  componentDidMount(): void {
    this.getMyToDos();
  }

  public getMyToDos = () => {
    /* console.log("test ToDo");
  alert("My ToDo") */ 

    // below code is default code to get
    this.props.context.msGraphClientFactory //default syntax
      .getClient("3") //version updated to 3
      .then((client: MSGraphClientV3): void => {
        client
          .api(
            "me/todo/lists/AQMkAGRiZGQAOTE5My05ZDRmLTRhNDYtYjMxZC01ZDUwOGQ4NTNmMTUALgAAA_tYM-SWNg9Lp0elhzgQ7O8BAJYyHGEAdpdJmI-9f86VAgMAAAIBEgAAAA==/tasks"
          ) //to get messages
          .version("v1.0")
          // selected columns from response preview
          .select("*")
          .get((err: any, res: any) => {
            this.setState({
              AllToDos: res.value,
            });
            // console.log(this.state.AllEmails);         //checking
            console.log(res);
            console.log(err);
          });
      });
  };

  public render(): React.ReactElement<IEventsProps> {
    return (
      <div  className={styles.main}
      style={{
        height: this.props.webHeight,
      }}
      >
        <h3>MyToDo list</h3>
        {this.state.AllToDos.map((todo) => {
          return (
            <div className={styles.Todo}>
              <p><b>{todo.title}</b></p>
              <p>{todo.status}</p>
              <p>{todo.importance}</p>
              <p>{todo.isReminderOn}</p>
              <p>{todo.categories}</p>
              <p>{moment(todo.createdDateTime).format("LL")}</p>
            {/*   <p>{todo.completedDateTime}</p>  */}
              <p>{todo.hasAttachments}</p> 
            </div>
          );
        })}
      </div>
    );
  }
}
