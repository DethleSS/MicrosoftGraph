import * as React from 'react';
import styles from './GraphWebPart.module.scss';
import { IGraphWebPartProps } from './IGraphWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from '@microsoft/sp-http'
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';
import { Web } from "@pnp/sp/presets/all";
import { ClientsidePageFromFile } from "@pnp/sp/clientside-pages";
import { sp } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/items';

export const GraphWebPart: React.FC<IGraphWebPartProps> = ({ context, contextGraphApi }) => {

  const [infoUser, setInfoUser] = React.useState<any>();
  const [infoGroup, setInfoGroup] = React.useState<any>();
  const [infoPlanner, setInfoPlanner] = React.useState<any>();
  const [userPhoto, setUserPhoto] = React.useState<any>();
  function getValueUser() {
    contextGraphApi.getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/me')
          .top(5)
          .get((error, infoUser: any, rawResponse?: any) => {
            setInfoUser(infoUser)
          });
      });
  };

  function getValuePhotoUser() {
    contextGraphApi.getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('me/photo/$value')
          .responseType('blob')
          .get()
          .then(data => { 
            const blobUrl = window.URL.createObjectURL(data)
            setUserPhoto(blobUrl)
          })
  })
}

  function getValueGroups() {
    contextGraphApi.getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/groups')
          .get((error, infoGroups: any, rawResponse?: any) => {
            setInfoGroup(infoGroups.value)
          });
      });
  };

  async function getValuePage() {
    let web = Web(context.pageContext.web.absoluteUrl + '/sites/TeamSite/');
    let page = await web.lists.getByTitle("Site Pages").items.get().then(el => {
      //console.log(el)
    })
    //return (page[1].CanvasContent1)
  }

  function getValuePlanner() {
    contextGraphApi.getClient()
      .then((client: MSGraphClient): void => {
        client
          .api('/me/planner/tasks')
          .get((error, _infoPlanner: any, rawResponse?: any) => {
            setInfoPlanner(_infoPlanner.value)
          });
      });
  };

  React.useEffect(() => {
    getValueUser();
    getValueGroups();
    getValuePlanner();
    getValuePhotoUser();
    getValuePage();
  }, [])

  return (
    <div className={styles.graphWebPart}>
      <div className={styles.container}>
        <div className={styles.row}>
          <div className={styles.column}>
            {infoGroup ?
              <>
                <h3>Group Name: {infoGroup[0].description}</h3>
                <h3>Visibility: {infoGroup[0].visibility}</h3>
                <h3>Creation Options Count: {infoGroup[0].creationOptions.length}</h3>
              </> : <h1>loading</h1>
            }
          </div>
          <div className={styles.column}>
            {infoPlanner ?
              <>
                <h3>Planner Task:</h3>
                {infoPlanner.length ?
                  <>
                    {infoPlanner.map((el: any) => <h3>{el.title}</h3>)}
                  </> : <h3>Tasks clear</h3>
                }
              </> : <h1>loading</h1>
            }
          </div>
          <div className={styles.column}>
            {infoUser ?
              <>
                {userPhoto ?
                  <img className={styles.logo__User} src={userPhoto} /> : <h3>...loading photo</h3>
                }
                <h3>Name: {infoUser.givenName}</h3>
                <h3>Surname: {infoUser.surname}</h3>
                <h3>Email: {infoUser.mail}</h3>
                <h3>Business Phones: {infoUser.businessPhones[0]}</h3>
              </> : <h1>loading</h1>
            }
          </div>
        </div>
      </div>
    </div>

  );
}
