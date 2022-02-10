/* eslint-disable @typescript-eslint/no-empty-function */
/* eslint-disable no-mixed-spaces-and-tabs */
/* eslint-disable no-tabs */
import { ITag } from "@fluentui/react";
import * as ReactDom from "react-dom";
import React = require("react");
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { ConnectionTagPicker } from "./ConnectionTagPicker";

export class PCFTagPicker implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private _container: HTMLDivElement;
  private _notifyOutputChanged: () => void;
  private _availableTags: ITag[];
  private _selectedTags: ITag[];

  //Variables for Properties
  private _tagsEntity: string;
  private _tagsEntityDisplayNameField: string;
  private _tagsEntityCategoryField: string;
  private _categoryEntity: string;
  private _categoryEntityNameField: string;
  private _categories: string;
  private _connectionRoleName: string;
  private _logString: string;

  // eslint-disable-next-line @typescript-eslint/no-empty-function
  constructor() {}
  updateView(context: ComponentFramework.Context<IInputs>): void {
      this.renderControl(context);
  }
  destroy(): void {
      ReactDom.unmountComponentAtNode(this._container);
  }
  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
      context: ComponentFramework.Context<IInputs>,
      notifyOutputChanged: () => void,
      state: ComponentFramework.Dictionary,
      container: HTMLDivElement,
  ): void {
      this._tagsEntity = context.parameters.tagsEntity.raw ?? "No Tags Entity Specified";
      this._tagsEntityDisplayNameField =
          context.parameters.tagsEntityDisplayNameField.raw ?? "No Tags Entity Display Name Specified";
      this._tagsEntityCategoryField =
          context.parameters.tagsEntityCategoryField.raw ?? "No Tags Entity Category Specified";

      this._categoryEntity = context.parameters.categoryEntity.raw ?? "No Category Entity Specified";
      this._categoryEntityNameField =
          context.parameters.categoryEntityNameField.raw ?? "No Category Entity Name Specified";
      this._categories = context.parameters.categories.raw ?? "";
      this._connectionRoleName =
          context.parameters.connectionRoleName.raw ?? "No Entity Relationship Names Specified";
      this._logString = `[TagPicker(${(<any>context).navigation._customControlProperties.controlId})]`;
      //Init local variables
      this._container = container;
      this._notifyOutputChanged = notifyOutputChanged;
      if ((<any>context).page.entityId == null) {
          return; //A new record is being created, no tags
      } else {
          this.getAvailableTags(context)
              .then(() => this.getConnectedTags(context))
              .then(() => this.renderControl(context));
      }
  }
  private renderControl(context: ComponentFramework.Context<IInputs>) {
      const recordSelector = React.createElement(ConnectionTagPicker, {
          availableTags: this._availableTags,
          selectedTags: this._selectedTags,
          onChange: (items?: ITag[]) => {
              if (items != undefined) {
                  items.forEach((tag) => {
                      if (this._selectedTags.indexOf(tag) == -1) {
                          console.debug(`${this._logString} Connecting ${tag.name}`);
                          this.connect(context, this._tagsEntity, tag.key, this._connectionRoleName);
                          this._selectedTags.push(tag);
                      }
                  });
                  const tagsToRemove: ITag[] = [];
                  this._selectedTags.forEach((tag) => {
                      if (items.indexOf(tag) == -1) {
                          console.debug(`${this._logString} Disconnecting ${tag.name}`);
                          this.disconnect(context, tag.key);
                          tagsToRemove.push(tag);
                      }
                  });
                  tagsToRemove.forEach((tag) => {
                      const index = this._selectedTags.indexOf(tag);
                      if (index > -1) {
                          console.debug(`${this._logString} Removing ${tag.name} from selected tags`);
                          this._selectedTags.splice(index, 1);
                      }
                  });
                  this.renderControl(context);
                  this._notifyOutputChanged();
              }
          },
      });

      ReactDom.render(recordSelector, this._container);
  }
  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
      return {
          tags: this._selectedTags
              .map(function (e) {
                  return e.name;
              })
              .join(", "),
      };
  }
  private getConnectedTags(context: ComponentFramework.Context<IInputs>) {
      const localContext = <any>context;
      const entityName = localContext.page.entityTypeName;
      return new Promise<void>((resolve, reject) => {
          context.utils.getEntityMetadata(this._tagsEntity).then((tagsMetadata) => {
              context.utils.getEntityMetadata(entityName).then((entityMetadata) => {
                  const entityIdFieldName = entityMetadata.PrimaryIdAttribute;
                  const entityId = localContext.page.entityId;
                  const tagsEntityIdFieldName = tagsMetadata.PrimaryIdAttribute;
                  const tagsEntityNameFieldName = this._tagsEntityDisplayNameField;
                  console.debug(`${this._logString} Getting connected tags`);

                  const fetchXml = `
                      <fetch>
                          <entity name='${this._tagsEntity}'>
                              <link-entity name='connection' from='record2id' to='${tagsEntityIdFieldName}'>
                                  <link-entity name='${entityName}' from='${entityIdFieldName}' to='record1id'>
                                      <filter>
                                          <condition attribute='${entityIdFieldName}' operator='eq' value='${entityId}'/>
                                      </filter>
                                  </link-entity>
                              </link-entity>
                          </entity>
                      </fetch>`;

                  const encodedFetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
                  context.webAPI.retrieveMultipleRecords(this._tagsEntity, encodedFetchXml).then(
                      (result) => {
                          this._selectedTags = result.entities.map((r) => {
                              return {
                                  key: r[tagsEntityIdFieldName],
                                  name: r[tagsEntityNameFieldName] ?? "Display Name is not available",
                              };
                          });
                          resolve();
                      },
                      (error) => {
                          console.error(`${this._logString} ${error}`);
                          reject();
                      },
                  );
              });
          });
      });
  }
  private getAvailableTags(context: ComponentFramework.Context<IInputs>) {
      return new Promise<void>((resolve, reject) => {
          context.utils.getEntityMetadata(this._tagsEntity).then(
              (metadata) => {
                  const entityIdFieldName = metadata.PrimaryIdAttribute;
                  const entityNameFieldName = this._tagsEntityDisplayNameField;
                  console.debug(`${this._logString} Getting tags`);
                  const categories = this._categories.split(",");
                  let fetchXml = "";
                  if (this._categories != "" && categories.length > 0) {
                      let filterString = "";
                      categories.forEach(function (item) {
                          filterString += "<value>" + item.trim() + "</value>";
                      });
                      fetchXml = `
                          <fetch>
                            <entity name='${this._tagsEntity}'>
                              <attribute name='${this._tagsEntityDisplayNameField}' />
                              <link-entity name='${this._categoryEntity}' from='${this._categoryEntity}id' to='${this._tagsEntityCategoryField}'>
                                <filter>
                                  <condition attribute='${this._categoryEntityNameField}' operator='in'>
                                  ${filterString}
                                  </condition>
                                </filter>
                              </link-entity>
                            </entity>
                          </fetch>`;
                  } else {
                      fetchXml = `
                          <fetch>
                            <entity name='${this._tagsEntity}'>
                              <attribute name='${this._tagsEntityDisplayNameField}' />
                            </entity>
                          </fetch>`;
                  }
                  const encodedFetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
                  context.webAPI.retrieveMultipleRecords(this._tagsEntity, encodedFetchXml).then((result) => {
                      this._availableTags = result.entities.map((r) => {
                          return {
                              key: r[entityIdFieldName],
                              name: r[entityNameFieldName] ?? "Display Name is not available",
                          };
                      });
                      resolve();
                  });
              },
              (error) => {
                  console.error(`${this._logString} ${error}`);
                  reject();
              },
          );
      });
  }
  private connect(
      context: ComponentFramework.Context<IInputs>,
      relatedEntityName: string,
      relatedEntityId: string | number,
      connectionRoleName: string,
  ) {
      const localContext = <any>context;
      const primaryEntityId = localContext.page.entityId;
      const primaryEntityName = localContext.page.entityTypeName;
      const clientUrl = localContext.page.getClientUrl();
      context.utils.getEntityMetadata(primaryEntityName).then((primaryMetadata) => {
          context.utils.getEntityMetadata(relatedEntityName).then((relatedMetadata) => {
              this.getConnectionRoleId(localContext, connectionRoleName).then((connectionRoleId) => {
                  const primaryCollectionName = primaryMetadata.EntitySetName;
                  const relatedCollectionName = relatedMetadata.EntitySetName;

                  const connection: any = {};
                  connection[
                      `record1id_${primaryEntityName}@odata.bind`
                  ] = `/${primaryCollectionName}(${primaryEntityId})`;
                  connection[
                      `record2id_${relatedEntityName}@odata.bind`
                  ] = `/${relatedCollectionName}(${relatedEntityId})`;
                  connection["record2roleid@odata.bind"] = `/connectionroles(${connectionRoleId})`;
                  connection["effectiveend"] = new Date();

                  const req = new XMLHttpRequest();
                  const localLogString = this._logString;
                  req.open("POST", `${clientUrl}/api/data/v9.0/connections`, true);
                  req.setRequestHeader("Accept", "application/json");
                  req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
                  req.setRequestHeader("OData-MaxVersion", "4.0");
                  req.setRequestHeader("OData-Version", "4.0");
                  req.onreadystatechange = function () {
                      if (this.readyState === 4) {
                          req.onreadystatechange = null;
                          if (this.status === 204 || this.status === 1223) {
                              //Success
                          } else {
                              console.error(`${localLogString} ${this.statusText}`);
                          }
                      }
                  };
                  req.send(JSON.stringify(connection));
              });
          });
      });
  }
  private disconnect(context: ComponentFramework.Context<IInputs>, relatedEntityId: string | number) {
      const localContext = <any>context;
      const primaryEntityId = localContext.page.entityId;
      const clientUrl = localContext.page.getClientUrl();
      this.getConnectionId(localContext, primaryEntityId, relatedEntityId).then((connectionId) => {
          const req = new XMLHttpRequest();
          req.open("DELETE", `${clientUrl}/api/data/v9.0/connections(${connectionId})`, false);

          const localLogString = this._logString;
          req.setRequestHeader("Accept", "application/json");
          req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
          req.setRequestHeader("OData-MaxVersion", "4.0");
          req.setRequestHeader("OData-Version", "4.0");
          req.onreadystatechange = function () {
              if (this.readyState === 4) {
                  req.onreadystatechange = null;
                  if (this.status === 204 || this.status === 1223) {
                      //Success
                  } else {
                      console.error(`${localLogString} ${this.statusText}`);
                  }
              }
          };
          req.send();
      });
  }
  private getConnectionRoleId(context: ComponentFramework.Context<IInputs>, connectionRoleName: string) {
      return new Promise<void>((resolve, reject) => {
          const fetchXml = `
          <fetch>
              <entity name="connectionrole">
                  <attribute name="connectionroleid" />
                  <filter>
                      <condition attribute="name" operator="eq" value="${connectionRoleName}" />
                  </filter>
              </entity>
          </fetch>`;
          const encodedFetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
          context.webAPI.retrieveMultipleRecords("connectionrole", encodedFetchXml).then(
              (result) => {
                  resolve(result.entities[0]["connectionroleid"]);
              },
              (error) => {
                  console.error(`${this._logString} ${error}`);
                  reject();
              },
          );
      });
  }
  private getConnectionId(
      context: ComponentFramework.Context<IInputs>,
      primaryEntityId: string,
      relatedEntityId: string | number,
  ) {
      return new Promise<void>((resolve, reject) => {
          const fetchXml = `
          <fetch>
              <entity name="connection">
                  <attribute name="connectionid" />
                  <filter>
                  <condition attribute="record1id" operator="eq" value="${primaryEntityId}" />
                  <condition attribute="record2id" operator="eq" value="${relatedEntityId}" />
                  </filter>
              </entity>
          </fetch>`;
          const encodedFetchXml = "?fetchXml=" + encodeURIComponent(fetchXml);
          context.webAPI.retrieveMultipleRecords("connection", encodedFetchXml).then(
              (result) => {
                  resolve(result.entities[0]["connectionid"]);
              },
              (error) => {
                  console.error(`${this._logString} ${error}`);
                  reject();
              },
          );
      });
  }
}
