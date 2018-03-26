require('./QueryDropdown.scss');

import * as React from 'react';
import { Dropdown, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { ConfigError } from '../../model/ApplicationError';
import { ErrorDisplay } from '../error/ErrorDisplay';

export interface IQueryDropdownProps {
  className?: string;
  stateKey?: string;
  label?: string;
  defaultKey?: string | number;
  isDisabled?: boolean;
  loadOptionsDelayMilliSecs?: number;
  onChanged?: (oldKey: string | number, newKey: string | number) => void;
  actions: {
    loadOptions: () => Promise<IDropdownOption[]>;
  };
}

export interface IQueryDropdownState {
  isLoading: boolean;
  options: IDropdownOption[];
  selectedIndex: number;
  error: Error;
}

export class QueryDropdown extends React.Component<IQueryDropdownProps, IQueryDropdownState> {
  private static readonly baseCssClassName = 's32-spfx-query-dropdown-property';
  private static readonly defaultLoadOptionsDelayMilliSecs = 1000;

  private loadOptionsDelayTimerId: number = null;

  constructor(props: IQueryDropdownProps) {
    super(props);

    this.onSelectionChanged = this.onSelectionChanged.bind(this);

    this.state = {
      isLoading: false,
      options: null,
      selectedIndex: null,
      error: null
    };
  }

  public componentDidMount() {
    this.loadOptions(this.props);
  }

  public componentWillReceiveProps(nextProps: IQueryDropdownProps) {
    if (nextProps.stateKey !== this.props.stateKey) {
      // Every change to the state key (this usually means every single key stroke) causes a change event, so delay the loading of options to allow time for the user to completely update the state key.
      window.clearTimeout(this.loadOptionsDelayTimerId);
      this.loadOptionsDelayTimerId = window.setTimeout(() => { this.loadOptions(nextProps); }, this.props.loadOptionsDelayMilliSecs ? this.props.loadOptionsDelayMilliSecs : QueryDropdown.defaultLoadOptionsDelayMilliSecs);
    }
    if (nextProps.defaultKey !== this.props.defaultKey) {
      // For non-reactive web parts, this resets the property to the last committed value.
      const newSelectedIndex = this.getOptionIndex(this.state.options, nextProps.defaultKey);
      if (newSelectedIndex !== this.state.selectedIndex) {
        this.setState({ selectedIndex: newSelectedIndex });
      }
    }
  }

  public render() {
    const className = QueryDropdown.baseCssClassName + (this.props.className ? ` ${this.props.className}` : '');

    let loadingMarkup = null;
    if (this.state.isLoading) {
      loadingMarkup =
        <Spinner size={SpinnerSize.small} label="Loading options..." ariaLive="assertive" />;
    }

    let errorMarkup = null;
    if (this.state.error) {
      errorMarkup =
        <ErrorDisplay
          className="ms-u-slideDownIn20"
          title="An error occurred while loading the options"
          actionMessage="Please try again and if the problem persists, please contact your support team."
          error={this.state.error}
        />;
    }

    const markup =
      <div className={className}>
        <Dropdown
          label={this.props.label}
          isDisabled={this.props.isDisabled || !this.state.options || this.state.options.length === 0}
          onChanged={this.onSelectionChanged}
          selectedKey={this.getOptionKey(this.state.selectedIndex, this.state.options)}
          options={this.state.options ? this.state.options : []} />
        {loadingMarkup}
        {errorMarkup}
      </div>;

    return markup;
  }

  protected loadOptions(props: IQueryDropdownProps) {
    if (props.actions && props.actions.loadOptions) {
      let oldSelectedKey: string | number = null;
      if (this.props.onChanged) {
        oldSelectedKey = this.getOptionKey(this.state.selectedIndex, this.state.options);
      }
      this.setState({
        isLoading: true,
        options: null,
        selectedIndex: null,
        error: null
      });
      props.actions.loadOptions()
        .then((options: IDropdownOption[]) => {
          const selectedIndex = this.getOptionIndex(options, this.props.defaultKey);
          this.setState({
            isLoading: false,
            options,
            selectedIndex
          });
          if (this.props.onChanged) {
            const newSelectedKey = this.getOptionKey(selectedIndex, options);
            this.props.onChanged(oldSelectedKey, newSelectedKey);
          }
        })
        .catch(error => {
          this.setState({
            isLoading: false,
            error
          });
        });
    }
    else {
      this.setState({ error: new ConfigError('Please update the configuration with the information required to retrieve link details.') });
    }
  }

  protected getOptionIndex(options: IDropdownOption[], optionKey?: string | number): number {
    let optionIndex = null;
    if (options && options.length > 0) {
      optionIndex = 0;
      if (optionKey) {
        let optionKeyIndex = null;
        for (let index = 0, length = options.length; index < length; index++) {
          const option = options[index];
          if (option.key === optionKey) {
            optionKeyIndex = index;
            break;
          }
        }
        if (optionKeyIndex !== null) {
          optionIndex = optionKeyIndex;
        }
      }
    }

    return optionIndex;
  }

  protected getOptionKey(index: number, options: IDropdownOption[]): string | number {
    return index !== null ? options[index].key : null;
  }

  protected onSelectionChanged(option: IDropdownOption, index: number) {
    if (this.props.onChanged) {
      this.props.onChanged(this.getOptionKey(this.state.selectedIndex, this.state.options), option.key);
    }
    this.setState({ selectedIndex: index });
  }
}
