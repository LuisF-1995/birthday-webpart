import * as React from 'react';
import './BirthDayCard.css';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IUserBirthdate } from '../../../models/IUserBirthdate';
import { PNP } from '../../../../services/Util';

interface IBirthDayCardStates{
    birthdayIcon:string;
}

export interface IBirthDayCardProps{
    context: WebPartContext;
    userInfo: IUserBirthdate;
}

export default class BirthDayCard extends React.Component<IBirthDayCardProps, IBirthDayCardStates> {

    private pnp:PNP;

    constructor(props:IBirthDayCardProps){
        super(props);
        this.pnp = new PNP(this.props.context);

        this.state = {
            birthdayIcon: ''
        }
    }

    componentDidMount(): void {
        this.getBirthdayIcon();
    }

    private getBirthdayIcon():void {
        this.pnp.getConfigValue('BirthdayIcon')
        .then(icon => {
            this.setState({
                birthdayIcon: icon
            })
        })
        .catch(error => {
            console.error(`Error al obtener el icono de cumpleaños. Error: ${error}`);
        })
    }

    public render(): React.ReactElement<IBirthDayCardProps> {
        const { Name, Birthdate } = this.props.userInfo;
        const formatter = new Intl.DateTimeFormat('es-CO', { month: 'long' });
        const formattedMonth = formatter.format(Birthdate);
        
        return (
            <section className='birth-card-container'>
                <div className='birth-icon-container'>
                    <img src={this.state.birthdayIcon} alt="Birthday Icon"/>
                </div>
                <article className='birth-card-info'>
                    <h3>
                        {Name}
                    </h3>
                    <p>
                        {`${Birthdate.getDate()} ${formattedMonth}`}
                    </p>
                </article>
            </section>
        );
    }
}
