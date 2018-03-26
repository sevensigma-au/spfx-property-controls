require('./ErrorDisplay.scss');
import * as React from 'react';

export interface IErrorDisplayProps {
  className?: string;
  title: string;
  actionMessage: string;
  error: Error;
}

export const ErrorDisplay = (props: IErrorDisplayProps): React.ReactElement<any> => {
  const className = 's32-error' + (props.className ? ` ${props.className}` : '');
  const markup =
    <div className={className}>
      <div className="s32-error-title">
        {props.title}
      </div>
      <div className="s32-error-action-message">
        {props.actionMessage}
      </div>
      <div className="s32-error-detail">
        Error details: {props.error.message}
      </div>
    </div>
    ;

  return markup;
};
