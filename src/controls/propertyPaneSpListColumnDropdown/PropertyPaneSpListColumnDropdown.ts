import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
  IPropertyPaneCustomFieldProps
} from '@microsoft/sp-webpart-base';
import { IDropdownOption } from 'office-ui-fabric-react/lib/DropDown';
import { WebStorageCache, StorageType } from '@sevensigma/web-data-store';
import { ConfigError } from '../../model/ApplicationError';
import { IKeyValuePair } from '../../model/BaseTypes';
import { SharePointDataService } from '../../services/SharePointDataService';
import { QueryDropdown, IQueryDropdownProps } from '../../common/queryDropdown/QueryDropdown';

export interface IPropertyPaneSpListColumnDropDownProps {
  webAbsoluteUrl: string;
  listTitle: string;
  defaultKey?: string;
  label?: string;
  disabled?: boolean;
  includeInternalColumns?: boolean;
  loadOptionsDelayMilliSecs?: number;
  cacheTimeoutSecs?: number;
  onPropertyChange?: (propertyPath: string, oldValue: any, newValue: any) => void;
}

export interface IPropertyPaneSpListColumnDropDownInternalProps extends IPropertyPaneSpListColumnDropDownProps, IPropertyPaneCustomFieldProps {}

export class PropertyPaneSpListColumnDropdownControl implements IPropertyPaneField<IPropertyPaneSpListColumnDropDownProps> {
  private static readonly baseComponentKey = 's32-spfx-splistcolumn-dropdown';
  private static readonly defaultCacheTimeoutSecs = 10;

  private dataService: SharePointDataService;
  private domElement: HTMLElement = null;
  private propertyPaneChangeCallback: (targetProperty?: string, newValue?: any) => void = null;

  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneSpListColumnDropDownInternalProps;

  constructor(targetProperty: string, properties: IPropertyPaneSpListColumnDropDownProps) {
    this.loadListTitles = this.loadListTitles.bind(this);
    this.onRender = this.onRender.bind(this);
    this.onChanged = this.onChanged.bind(this);

    this.targetProperty = targetProperty;
    const cacheTimeoutSecs = properties.cacheTimeoutSecs ? properties.cacheTimeoutSecs : PropertyPaneSpListColumnDropdownControl.defaultCacheTimeoutSecs;
    const cache = new WebStorageCache(PropertyPaneSpListColumnDropdownControl.baseComponentKey, cacheTimeoutSecs, StorageType.sessionStorage);
    this.dataService = new SharePointDataService(cache);

    this.properties = {
      key: this.getComponentKey(targetProperty),
      webAbsoluteUrl: properties.webAbsoluteUrl,
      listTitle: properties.listTitle,
      defaultKey: properties.defaultKey,
      label: properties.label,
      disabled: properties.disabled,
      includeInternalColumns: properties.includeInternalColumns,
      loadOptionsDelayMilliSecs: properties.loadOptionsDelayMilliSecs,
      cacheTimeoutSecs: properties.cacheTimeoutSecs,
      onPropertyChange: properties.onPropertyChange,
      onRender: this.onRender
    };
  }

  public render() {
    if (this.domElement) {
      this.onRender(this.domElement);
    }
  }

  protected getComponentKey(targetProperty: string) {
    let key = PropertyPaneSpListColumnDropdownControl.baseComponentKey;
    if (targetProperty) {
      key += `-${targetProperty.replace(' ', '-')}`;
    }
    return key;
  }

  protected onRender(domElement: HTMLElement, context?: any, changeCallback?: (targetProperty?: string, newValue?: any) => void) {
    if (!this.domElement) {
      this.domElement = domElement;
    }
    this.propertyPaneChangeCallback = changeCallback;
    const element: React.ReactElement<IQueryDropdownProps> = React.createElement(QueryDropdown, {
      label: this.properties.label,
      stateKey: `${this.properties.webAbsoluteUrl}-${this.properties.listTitle}`,
      defaultKey: this.properties.defaultKey,
      isDisabled: this.properties.disabled,
      loadOptionsDelayMilliSecs: this.properties.loadOptionsDelayMilliSecs,
      onChanged: this.onChanged,
      actions: {
        loadOptions: this.loadListTitles
      }
    });
    ReactDom.render(element, domElement);
  }

  protected loadListTitles(): Promise<IDropdownOption[]> {
    return new Promise<IDropdownOption[]>(
      (resolve, reject) => {
        try {
          if (this.properties.webAbsoluteUrl && this.properties.listTitle) {
            this.dataService.getListColumns(this.properties.webAbsoluteUrl, this.properties.listTitle, this.properties.includeInternalColumns)
              .then((listColumns: IKeyValuePair<string, string>[]) => {
                const options = listColumns.map(
                  (listColumn: IKeyValuePair<string, string>) => {
                    return {
                      key: listColumn.key,
                      text: `${listColumn.value} (${listColumn.key})`
                    };
                  }
                );
                resolve(options);
              })
              .catch(error => {
                reject(error);
              });
          }
          else {
            reject(new ConfigError('The URL and/or the list title wasn\'t provided.'));
          }
        }
        catch (error) {
          reject(error);
        }
      }
    );
  }

  protected onChanged(oldKey: string | number, newKey: string | number) {
    // Notify the property pane of the change so that it can handle it appropariately e.g. temporarily store values for non-reactive property panels until Apply is clicked.
    if (this.propertyPaneChangeCallback) {
      this.propertyPaneChangeCallback(this.targetProperty, newKey);
    }
    // Allow caller customisation of the onPropertyChange event.
    if (this.properties.onPropertyChange) {
      this.properties.onPropertyChange(this.targetProperty, oldKey, newKey);
    }
  }
}

export const PropertyPaneSpListColumnDropdown = (targetProperty: string, properties: IPropertyPaneSpListColumnDropDownProps): IPropertyPaneField<IPropertyPaneSpListColumnDropDownProps> => {
  return new PropertyPaneSpListColumnDropdownControl(targetProperty, properties);
};