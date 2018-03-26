/* tslint:disable:no-misused-new */

export interface IApplicationError {
  name: string;
  message: string;
  new (message: string): IApplicationError;
}

export const ApplicationError = ((message: string): void => {
  this.name = 'ApplicationError';
  this.message = message;
}) as any as IApplicationError;

ApplicationError.prototype = Object.create(Error.prototype);
ApplicationError.prototype.constructor = ApplicationError;

export class ConfigError extends ApplicationError {
  constructor(message: string) {
    super(message);
    this.message = message;
  }
}
