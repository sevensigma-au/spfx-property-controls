import { sp, Web } from '@pnp/sp';
import { ErrorHelper } from '@sevensigma/pnpjs-error-helper';
import { WebStorageCache, SharedDataSynchroniser } from '@sevensigma/web-data-store';
import {
  ISpList,
  SpFieldType,
  ISpField,
  SpFieldProperty,
  SpListProperty
} from '../model/SharePointTypes';
import { IKeyValuePair } from '../model/BaseTypes';

export class SharePointDataService {
  private cache: WebStorageCache;
  private sharedDataSync: SharedDataSynchroniser;

  constructor(cache: WebStorageCache) {
    this.cache = cache;
    if (cache && cache.timeoutSecs > 0) {
      this.sharedDataSync = new SharedDataSynchroniser(cache);
    }
  }

  public getCustomListTitles(webAbsoluteUrl: string, sharedDataTimeoutSecs?: number): Promise<string[]> {
    return new Promise<string[]>(
      (resolve, reject) => {
        try {
          const dataCacheKey = `lists-${webAbsoluteUrl}`;
          let listTitles: string[];

          if (this.cache) {
            listTitles = this.cache.getItem(dataCacheKey);
          }

          if (listTitles) {
            resolve(listTitles);
          }
          else {
            if (this.sharedDataSync && this.sharedDataSync.isLoading(dataCacheKey)) {
              this.sharedDataSync.waitForData(dataCacheKey, sharedDataTimeoutSecs)
                .then((sharedListTitles: string[]) => {
                  if (sharedListTitles) {
                    resolve(sharedListTitles);
                  }
                  else {
                    reject(new Error('Unable to retrieve list titles from shared data.'));
                  }
                })
                .catch(error => {
                  reject(error);
                });
            }
            else {
              if (this.sharedDataSync) {
                this.sharedDataSync.setIsLoading(dataCacheKey, true);
              }
              const web = webAbsoluteUrl ? new Web(webAbsoluteUrl) : sp.web;
              web.lists
              .filter(`(BaseType eq 0) and (BaseTemplate eq 100) and (Title ne 'TaxonomyHiddenList')`)
              .orderBy(SpListProperty.title, true)
              .get()
                .then((lists: ISpList[]) => {
                  listTitles = lists.map((list: ISpList) => {
                    return list.Title;
                  });
                  if (this.sharedDataSync) {
                    this.sharedDataSync.setIsLoading(dataCacheKey, false);
                  }
                  if (this.cache) {
                    this.cache.setItem(dataCacheKey, listTitles);
                  }
                  resolve(listTitles);
                })
                .catch(error => {
                  if (this.sharedDataSync) {
                    this.sharedDataSync.setIsLoading(dataCacheKey, false);
                  }
                  reject(ErrorHelper.getErrorWithDetails(error));
                });
            }
          }
        }
        catch (error) {
          reject(error);
        }
      }
    );
  }

  public getListColumns(webAbsoluteUrl: string, listTitle: string, includeInternalColumns: boolean = false, filter?: string, sharedDataTimeoutSecs?: number): Promise<IKeyValuePair<string, string>[]> {
    return new Promise<IKeyValuePair<string, string>[]>(
      (resolve, reject) => {
        try {
          const dataCacheKey = `columns-${webAbsoluteUrl}-${listTitle}-${includeInternalColumns ? 'with-internal' : 'no-intenal'}`;
          let listColumns: IKeyValuePair<string, string>[];

          if (this.cache) {
            listColumns = this.cache.getItem(dataCacheKey);
          }

          if (listColumns) {
            resolve(listColumns);
          }
          else {
            if (this.sharedDataSync && this.sharedDataSync.isLoading(dataCacheKey)) {
              this.sharedDataSync.waitForData(dataCacheKey, sharedDataTimeoutSecs)
                .then((sharedListColumns: IKeyValuePair<string, string>[]) => {
                  if (sharedListColumns) {
                    resolve(sharedListColumns);
                  }
                  else {
                    reject(new Error('Unable to retrieve list columns from shared data.'));
                  }
                })
                .catch(error => {
                  reject(error);
                });
            }
            else {
              if (this.sharedDataSync) {
                this.sharedDataSync.setIsLoading(dataCacheKey, true);
              }
              const web = webAbsoluteUrl ? new Web(webAbsoluteUrl) : sp.web;
              const internalColumnFilter: string = `${SpFieldProperty.internalName} eq 'Title'
                or ${SpFieldProperty.fieldTypeKind} ne ${SpFieldType.contentTypeId}
                and ${SpFieldProperty.fieldTypeKind} ne ${SpFieldType.guid}
                and ${SpFieldProperty.hidden} eq false
                and ${SpFieldProperty.canBeDeleted} eq true`;
              web.lists.getByTitle(listTitle).fields
              .select(SpFieldProperty.title, SpFieldProperty.internalName)
              .filter(`Title ne ''` + (filter ? ` and ${filter}` : '') + (!includeInternalColumns ? ` and ${internalColumnFilter}` : ''))
              .orderBy(SpFieldProperty.title, true)
              .get()
                .then((columns: ISpField[]) => {
                  listColumns = [];
                  for (let index = 0, length = columns.length; index < length; index++) {
                    const column = columns[index];
                    if (!includeInternalColumns && column.InternalName.substr(1, 1) === '_') {
                      // Ignore internal column not caught by REST API filter.
                    }
                    else {
                      listColumns.push({
                        key: column.InternalName,
                        value: column.Title
                      });
                    }
                  }
                  if (this.sharedDataSync) {
                    this.sharedDataSync.setIsLoading(dataCacheKey, false);
                  }
                  if (this.cache) {
                    this.cache.setItem(dataCacheKey, listColumns);
                  }
                  resolve(listColumns);
                })
                .catch(error => {
                  if (this.sharedDataSync) {
                    this.sharedDataSync.setIsLoading(dataCacheKey, false);
                  }
                  reject(ErrorHelper.getErrorWithDetails(error));
                });
            }
          }
        }
        catch (error) {
          reject(error);
        }
      }
    );
  }
}
