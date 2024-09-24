import * as React from 'react';
import styles from './BirthDay.module.scss';
import type { IBirthDayProps } from './IBirthDayProps';
import BirthDayCard from './birthDayCard/BirthDayCard';
import { PNP } from '../../../services/Util';
import { ISiteUserInfo } from '@pnp/sp/presets/all';
import Swal from 'sweetalert2';
import { IUserBirthdate } from '../../models/IUserBirthdate';

interface IBirthdayStates {
  birthdayUsers:IUserBirthdate[];
  birthdayGroupName:string;
}

export default class BirthDay extends React.Component<IBirthDayProps, IBirthdayStates> {
  private pnp:PNP;

  constructor(props:IBirthDayProps){
    super(props);
    this.pnp = new PNP(this.props.context);

    this.state = {
      birthdayGroupName:'',
      birthdayUsers: [],
    };
  }
  async componentDidMount(): Promise<void> {
    const birthDayGroupNameConfigTitle = 'BirthdayGroupName';
    await this.addUserToGroup(birthDayGroupNameConfigTitle);
    await this.loadBirthdayUsers(birthDayGroupNameConfigTitle);
  }

  private addUserToGroup = async (groupItemTitle: string): Promise<void> => {
    try {
      const groupname:string = await this.pnp.getConfigValue(groupItemTitle);
      const groupUser = await this.pnp.addUserToGroup(groupname);
      
      if(groupUser){
        const currentUserInfo:ISiteUserInfo = await this.pnp.getCurrentUserInfo();
        const currentUserBirthdate:string = await this.pnp.getUserBirthDate(currentUserInfo.LoginName);

        if(currentUserBirthdate.length === 0)
          Swal.fire({
            title:'Registrar fecha de cumpleaños',
            text: '',
            icon: 'info',
            input: 'date',
            showLoaderOnConfirm: true,
            preConfirm: async (birthdate:string) => {
              try {
                const birthdateRegistered = await this.pnp.setUserBirthDate(currentUserInfo.LoginName, birthdate);
                if(birthdateRegistered)
                  await Swal.fire('Fecha de cumpleaños registrada exitosamente', '', 'success');
                return birthdateRegistered;
              } catch (error) {
                return Swal.showValidationMessage(`Error al registrar la fecha de cumpleaños, por favor vuelve a intentarlo mas tarde`);
              }
            },
            allowOutsideClick: () => !Swal.isLoading(),
            confirmButtonText: 'Guardar'
          })
          .then(async (result) => {
            if(result.isConfirmed){
              await this.loadBirthdayUsers(groupItemTitle);
              Swal.close();
            }
          })
          .catch(error => {
            Swal.close();
          })
      }
    } catch (error) {
      console.error(`Error al intentar agregar usuario al grupo. Error: ${error}`);
    }
  }

  private async loadBirthdayUsers(groupItemTitle: string): Promise<void> {
    try {
      const groupname:string = await this.pnp.getConfigValue(groupItemTitle);
      const siteGroupUsers:ISiteUserInfo[] = await this.pnp.getUsersFromGroup(groupname);
      const userBirthDateInfo:IUserBirthdate[] = [];

      for (const user of siteGroupUsers) {
        const birthdate = await this.pnp.getUserBirthDate(user.LoginName);
        const [day, month] = birthdate.split("/");
        const actualYear = new Date().getFullYear();

        const userBirthdate:IUserBirthdate = {
          Id: user.Id,
          LoginName: user.LoginName,
          Name: user.Title,
          Email: user.Email,
          Birthdate: new Date(actualYear, parseInt(month) - 1, parseInt(day)),
        };

        userBirthDateInfo.push(userBirthdate);
      }

      this.setState({
        birthdayUsers: userBirthDateInfo
      });

    } catch (error) {
      console.error(`Error al intentar obtener los usuarios que cumplen años. Error: ${error}`);
    }
  }

  public render(): React.ReactElement<IBirthDayProps> {
    const {
      hasTeamsContext,
    } = this.props;

    const {birthdayUsers} = this.state;

    return (
      <main className={`${styles.birthDay} ${hasTeamsContext ? styles.teams : ''}`}>
        <h2>Cumpleaños</h2>
        <div className={styles.birthContainer}>
          {
            birthdayUsers.length > 0 && 
            birthdayUsers.filter((userBirthInfo: IUserBirthdate) => userBirthInfo.Birthdate > new Date() && userBirthInfo.Birthdate.getMonth() === new Date().getMonth()).length > 0 ?
            birthdayUsers
            .sort((a: IUserBirthdate, b: IUserBirthdate) => new Date(a.Birthdate).getTime() - new Date(b.Birthdate).getTime())
            .map((userBirthInfo:IUserBirthdate, index:number) => {
              if (userBirthInfo.Birthdate > new Date())
                return(
                  <BirthDayCard key={index} context={this.props.context} userInfo={userBirthInfo} />
                )
            })
            :
            <h4>No hay cumpleañeros este mes</h4>
          }
        </div>
      </main>
    );
  }
}
