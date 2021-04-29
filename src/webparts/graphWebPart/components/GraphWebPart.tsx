import * as React from 'react';
import styles from './GraphWebPart.module.scss';
import { IGraphWebPartProps } from './IGraphWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-http'
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export const GraphWebPart: React.FC<IGraphWebPartProps> = ({ context }) => {

  const [infoUser, setInfoUser] = React.useState<any>();
  const [infoGroup, setInfoGroup] = React.useState<any>();
  const [infoPlanner, setInfoPlanner] = React.useState<any>();

  function getValueUser() {
    context.getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/me')
          .top(5)
          .get((error, infoUser: any, rawResponse?: any) => {
            console.log(infoUser)
            setInfoUser(infoUser)
          });
      });
  };

  function getValueGroups() {
    context.getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/groups')
          .get((error, infoGroups: any, rawResponse?: any) => {
            console.log(infoGroups)
            setInfoGroup(infoGroups.value)
          });
      });
  };

  function getValuePlanner() {
    context.getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/me/planner/tasks')
          .get((error, _infoPlanner: any, rawResponse?: any) => {
            console.log(_infoPlanner)
            setInfoPlanner(_infoPlanner.value)
          });
      });
  };

  React.useEffect(() => {
    getValueUser();
    getValueGroups();
    getValuePlanner();
  }, [])

  return (
    <div className={styles.graphWebPart}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            {infoGroup ?
              <>
                <h1>Group Name: {infoGroup[0].description}</h1>
                <h1>Visibility: {infoGroup[0].visibility}</h1>
                <h1>Creation Options Count: {infoGroup[0].creationOptions.length}</h1>
              </> : <h1>loading</h1>
            }
          </div>
          <div className={styles.column}>
            {infoPlanner ?
              <>
                <h1>Planner Task:</h1>
                {infoPlanner.length ?
                  <>
                    {infoPlanner.map((el: any) => <h1>{el}</h1>)}
                  </> : <h1>Tasks clear</h1>
                }
              </> : <h1>loading</h1>
            }
          </div>
          <div className={styles.column}>
            {infoUser ?
              <>
                <h1>Name: {infoUser.givenName}</h1>
                <h1>Surname: {infoUser.surname}</h1>
                <h1>Email: {infoUser.mail}</h1>
                <h1>Business Phones: {infoUser.businessPhones[0]}</h1>
              </> : <h1>loading</h1>
            }
          </div>
        </div>
      </div>
    </div>

  );
}
