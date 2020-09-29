import * as React from 'react';
import './statusTaskModule.scss';
import { getSentNotification } from '../../apis/messageListApi';
import { RouteComponentProps } from 'react-router-dom';
import * as AdaptiveCards from "adaptivecards";
import { Loader, Flex, FlexItem, Menu, MenuItem, Text, Button } from '@stardust-ui/react';
import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';
import { Icon as IconFabric } from 'office-ui-fabric-react/lib/Icon';
import {
    getInitAdaptiveCard, setCardTitle, setCardImageLink, setCardSummary,
    setCardAuthor, setCardBtn
} from '../AdaptiveCard/adaptiveCard';

const repeatFrequencyValues = ['D\u00eda', 'Semanal', 'Mes'];

export interface IMessage {
    id: string;
    title: string;
    acknowledgements?: string;
    reactions?: string;
    responses?: string;
    succeeded?: string;
    failed?: string;
    throttled?: string;
    acknowledged?: string;
    isRecurrence?: boolean;
    sentDate?: string;
    imageLink?: string;
    summary?: string;
    author?: string;
    buttonLink?: string;
    buttonTitle?: string;
    buttonLink2?: string;
    buttonTitle2?: string;
    teamNames?: string[];
    rosterNames?: string[];
    allUsers?: boolean;
    sendingStartedDate?: string;
    sendingDuration?: string;
    repeats?: string,
    repeatFor?: number,
    repeatFrequency?: string,
    weekSelection?: string,
    repeatStartDate?: string,
    repeatEndDate?: string,
}

export interface IStatusState {
    message: IMessage;
    loader: boolean;
}

class StatusTaskModule extends React.Component<RouteComponentProps, IStatusState> {
    private initMessage = {
        id: "",
        title: ""
    };

    private card: any;

    constructor(props: RouteComponentProps) {
        super(props);

        this.card = getInitAdaptiveCard();

        this.state = {
            message: this.initMessage,
            loader: true
        };
    }

    public componentDidMount() {
        let params = this.props.match.params;
        console.log(this.props.match.params);

        if ('id' in params) {
            let id = params['id'];
            this.getItem(id).then(() => {
                this.setState({
                    loader: false
                }, () => {
                    setCardTitle(this.card, this.state.message.title);
                    setCardImageLink(this.card, this.state.message.imageLink);
                    setCardSummary(this.card, this.state.message.summary);
                    setCardAuthor(this.card, this.state.message.author);
                    setCardBtn(this.card, this.state.message.buttonTitle, this.state.message.buttonLink, this.state.message.buttonTitle2, this.state.message.buttonLink2);

                    let adaptiveCard = new AdaptiveCards.AdaptiveCard();
                    adaptiveCard.parse(this.card);
                    let renderedCard = adaptiveCard.render();
                    document.getElementsByClassName('adaptiveCardContainer')[0].appendChild(renderedCard);
                    const primaryButtonTitle = this.state.message.buttonTitle;
                    const primaryButtonLink = this.state.message.buttonLink;
                    const secondaryButtonLink = this.state.message.buttonLink2;
                    adaptiveCard.onExecuteAction = function (action) {
                        if (action.title === primaryButtonTitle) {
                            window.open(primaryButtonLink, '_blank');
                        }
                        else {
                            window.open(secondaryButtonLink, '_blank');
                        }
                    }
                });
            });
        }
    }

    private getItem = async (id: number) => {
        try {
            const response = await getSentNotification(id);
            response.data.sendingDuration = this.formatNotificationSendingDuration(response.data.sendingStartedDate, response.data.sentDate);
            response.data.sendingStartedDate = this.formatNotificationDate(response.data.sendingStartedDate);
            response.data.sentDate = this.formatNotificationDate(response.data.sentDate);

            this.setState({
                message: response.data
            });
        } catch (error) {
            return error;
        }
    }

    private formatNotificationSendingDuration = (sendingStartedDate: string, sentDate: string) => {
        let sendingDuration = "";
        if (sendingStartedDate && sentDate) {
            let timeDifference = new Date(sentDate).getTime() - new Date(sendingStartedDate).getTime();
            sendingDuration = new Date(timeDifference).toISOString().substr(11, 8);
        }
        return sendingDuration;
    }

    private formatNotificationDate = (notificationDate: string) => {
        if (notificationDate) {
            notificationDate = (new Date(notificationDate)).toLocaleString(navigator.language, { year: 'numeric', month: 'numeric', day: 'numeric', hour: 'numeric', minute: 'numeric', hour12: true });
            notificationDate = notificationDate.replace(',', '\xa0\xa0');
        }
        return notificationDate;
    }

    private openLink = (event: any) => {
        window.open(this.state.message.buttonLink2, '_blank');
    }
    public render(): JSX.Element {
        let menuItems: {}[] = [];
        menuItems.push(<MenuItem key="Sent" content="Enviados" > Enviados ({this.state.message.succeeded})
                        </MenuItem>);
        menuItems.push(<MenuItem key="Not delivered" content="No entregados" >No entregados ({this.state.message.failed})
        </MenuItem>);
        menuItems.push(<MenuItem key="Throttled" content="Reintentando">Reintentando ({this.state.message.throttled})
        </MenuItem>);
        menuItems.push(<MenuItem key="Acknowledged" content="Le&iacute;do">Le&iacute;do ({this.state.message.acknowledged})
        </MenuItem>);

        let recurrenceMessage: string = "";
        if (this.state.message.isRecurrence) {
            recurrenceMessage = "Occure ";
            let repeats: string = this.state.message.repeats ? this.state.message.repeats : "";
            let repeatFrequency: string = this.state.message.repeatFrequency ? this.state.message.repeatFrequency : "";

            if (this.state.message.repeats !== "Personalizar") {
                recurrenceMessage += repeats.toLowerCase();
            }
            else if (this.state.message.repeats === "Personalizar") {
                if (this.state.message.repeatFrequency === repeatFrequencyValues[0] || this.state.message.repeatFrequency === repeatFrequencyValues[2]) {
                    recurrenceMessage += "cada dia de la " + this.state.message.repeatFor + " " + repeatFrequency.toLowerCase();
                }
                else if (this.state.message.repeatFrequency === repeatFrequencyValues[1]) {
                    let weeks = "";
                    let weekSelection: string = this.state.message.weekSelection ? this.state.message.weekSelection : "";
                    if (weekSelection.indexOf('0') !== -1) {
                        weeks = "Monday,";
                    }
                    if (weekSelection.indexOf('1') !== -1) {
                        weeks += "Tuesday,";
                    }
                    if (weekSelection.indexOf('2') !== -1) {
                        weeks += "Wednesday,";
                    }
                    if (weekSelection.indexOf('3') !== -1) {
                        weeks += "Thursday,";
                    }
                    if (weekSelection.indexOf('4') !== -1) {
                        weeks += "Friday,";
                    }
                    if (weekSelection.indexOf('5') !== -1) {
                        weeks += "Saturday,";
                    }
                    if (weekSelection.indexOf('6') !== -1) {
                        weeks += "Sunday,";
                    }

                    recurrenceMessage += "cada dia de la " + this.state.message.repeatFor + " " + repeatFrequency.toLowerCase() + "(" + weeks.substring(0, weeks.length - 1) + ")";
                }
            }
            let repeatStartDate: string = this.formatNotificationDate(this.state.message.repeatStartDate ? this.state.message.repeatStartDate : "");
            recurrenceMessage += " comenzando " + repeatStartDate.substring(0, repeatStartDate.length - 8);
        }



        if (this.state.loader) {
            return (
                <div className="Loader">
                    <Loader />
                </div>
            );
        } else {
            return (
                <div className="taskModule">
                    <div className="formContainer">
                        <div className="formContentContainer" >
                            <div className="contentField">
                                <h3>Titulo</h3>
                                <span>{this.state.message.title}</span>
                            </div>
                            <div className="contentField">
                                <Flex>
                                    <FlexItem >
                                        <h3>Detalles del envio</h3>
                                    </FlexItem>
                                    <FlexItem push className="ml-auto">
                                        <h3>Duraci&oacute;n</h3>
                                    </FlexItem>
                                </Flex>
                                <Flex>
                                    <FlexItem >
                                        <span>{this.state.message.sendingStartedDate} --> {this.state.message.sentDate}</span>
                                    </FlexItem>
                                    <FlexItem push>
                                        <span>{this.state.message.sendingDuration}</span>
                                    </FlexItem>
                                </Flex>
                            </div>

                            <div className="contentField">
                                <h3>Mensaje recurrente</h3>
                                <Flex gap="gap.small">
                                    <FlexItem>
                                        <Text
                                            content={this.state.message.isRecurrence ? "Yes" : "No"} />
                                    </FlexItem>
                                    <FlexItem>
                                        <IconFabric iconName='Sync' />
                                    </FlexItem>
                                    <FlexItem>
                                        <Text
                                            content={recurrenceMessage} />
                                    </FlexItem>
                                </Flex>
                            </div>
                            <div className="contentField" hidden={!this.state.message.buttonTitle2}>
                                <Flex gap="gap.small">
                                    <FlexItem>
                                        <h3>{this.state.message.buttonTitle2}</h3>
                                    </FlexItem>
                                    <FlexItem>
                                        <Button className="openLinkBrn" text content="Abierto" onClick={this.openLink} />
                                    </FlexItem>
                                </Flex>
                            </div>
                            <div className="contentField">
                                <h3>Estado</h3>
                            </div>
                            <div className="contentField">
                                <Menu defaultActiveIndex={0} className="tab-menu" items={menuItems} primary />
                            </div>
                            <div className="contentField">
                                {this.renderAudienceSelection()}
                            </div>
                        </div>
                        <div className="adaptiveCardContainer">
                        </div>
                    </div>

                    <div className="footerContainer">
                        <div className="buttonContainer">
                        </div>
                    </div>
                </div>
            );
        }
    }

    private renderAudienceSelection = () => {
        if (this.state.message.teamNames && this.state.message.teamNames.length > 0) {
            let length = this.state.message.teamNames.length;
            return (
                <div>
                    <h3>Enviar al canal general de los equipos</h3>
                    {this.state.message.teamNames.sort().map((team, index) => {
                        if (length === index + 1) {
                            return (<span key={`teamName${index}`} >{team}</span>);
                        } else {
                            return (<span key={`teamName${index}`} >{team}, </span>);
                        }
                    })}
                </div>);
        } else if (this.state.message.rosterNames && this.state.message.rosterNames.length > 0) {
            let length = this.state.message.rosterNames.length;
            return (
                <div>
                    <h3>Enviado via chat a miembros de/los equipo(s)</h3>
                    {this.state.message.rosterNames.sort().map((team, index) => {
                        if (length === index + 1) {
                            return (<span key={`teamName${index}`} >{team}</span>);
                        } else {
                            return (<span key={`teamName${index}`} >{team}, </span>);
                        }
                    })}
                </div>);
        } else if (this.state.message.allUsers) {
            return (
                <div>
                    <h3>Enviar via chat a todos</h3>
                </div>);
        } else {
            return (<div></div>);
        }
    }
}

export default StatusTaskModule;