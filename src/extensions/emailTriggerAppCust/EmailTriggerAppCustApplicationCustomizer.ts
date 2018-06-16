import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import {SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import * as strings from 'EmailTriggerAppCustApplicationCustomizerStrings';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './EmailTriggerAppCustApplicationCustomizer.module.scss';
import * as jQuery from 'jquery';
const LOG_SOURCE: string = 'Email Trigger ApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IEmailTriggerAppCustApplicationCustomizerProperties {
  // This is an example; replace with your own property
  listGuid: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class EmailTriggerAppCustApplicationCustomizer
  extends BaseApplicationCustomizer<IEmailTriggerAppCustApplicationCustomizerProperties> {

   private _Bottom : PlaceholderContent  | undefined;

  private _validateUser() : Promise<boolean>{
    return new Promise<boolean>((resolve : (data : any) => void, reject : (error : any) => void) => {
        this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${this.properties.listGuid}')/items?$filter(UserEmail eq ${this.context.pageContext.user.email})`, SPHttpClient.configurations.v1,{
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        }).then((response : SPHttpClientResponse) => {
          return response.json();
        }).then((data : any) => {
            resolve(data.value.length > 0);
        }).catch((error: any) =>{
            reject(error);
            Log.info(LOG_SOURCE, error);
        });
    });
  }
  
  private _RegisterUserToSite() : Promise<boolean>{
      return new Promise<boolean>((resolve : (data : any) => void , reject : (error : any) => void) => {
          this._getListItemEntityTypeFullName().then((entityTypeFullName : string) : Promise<SPHttpClientResponse> => {
              if(entityTypeFullName !== "Error"){
                  const spData : string = JSON.stringify({
                    '__metadata' : {
                      'type' : entityTypeFullName
                    },
                    'Title': `${this.context.pageContext.user.displayName}`,
                    'UserEmail' : `${this.context.pageContext.user.email}`
                  });

                  return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${this.properties.listGuid}')/items`, SPHttpClient.configurations.v1, {
                    headers: {
                      'Accept': 'application/json;odata=nometadata',
                      'Content-type': 'application/json;odata=verbose',
                      'odata-version': ''
                    },
                    body : spData
                  });
              }
          }).then((resposne : SPHttpClientResponse) => {
              resolve(resposne.status);
              return;
          }).catch((error : any) => {
              console.log(error);
          });
          
      });
  }

  private _getListItemEntityTypeFullName() : Promise<string>{
      return new Promise<string>((resolve : (data : any) => void, reject : (error : any) => void) => {
          this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbyid('${this.properties.listGuid}')/ListItemEntityTypeFullName`, SPHttpClient.configurations.v1, {
            headers: {
              'Accept': 'application/json;odata=nometadata',
              'odata-version': ''
            }
          }).then((response : SPHttpClientResponse) => {
            return response.json();
          }).then((data : any) => {
            resolve(data.value);
            return;
          }).catch((error : any) => {
            reject('Error');
            console.log(error);
          });
      });
  }

  private _onDispose(): void {
    console.log('Disposed custom top and bottom placeholders.');
  }

  private _ShowRegistration() : void {
      if(!this._Bottom){
        this._Bottom = this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          {onDispose: this._onDispose}
        );
      }

      if(this._Bottom.domElement){
        this._Bottom.domElement.innerHTML = `<div id=${strings.DivId} style="line-height:4px;"><p>Thanks for visiting!</p><p>We have registered you in our site.</p></div>`;
      }

      let placeHolderDecoration = document.getElementById(strings.DivId);
      if(placeHolderDecoration != null && placeHolderDecoration != undefined){
        let parent = placeHolderDecoration.parentElement;
        console.log(parent);
        parent.className = `${styles.placeHolderArea} ${styles["fade-in"]} ${styles.one}`;
        
        jQuery('[class*=placeHolderArea').fadeOut(8000);
                
      }
  }

  @override
  public onInit(): Promise<void> {
    let checkIfExists : boolean; 
    console.log(escape(strings.Title));

    console.log('Current Logged In User: ' + escape(this.context.pageContext.user.displayName));

    console.log(`Logged In User Email: ${this.context.pageContext.user.email}`);

    this._validateUser().then((data : boolean) => {
        !data ? 
        this._RegisterUserToSite().then((isRegistered : boolean) => {
          isRegistered ? this._ShowRegistration() : console.log('Not able to add User')
        }).catch((error: any) => {
          console.log(error);
        })
        : console.log("User Already Exisits");
    });

    return Promise.resolve();
  }
}
