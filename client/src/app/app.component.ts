import { Component } from '@angular/core';
import { ExcelService } from './excel.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  constructor(private excelService: ExcelService) {
  }
  data: any;
  generateExcel() {
    this.data = [
      {
        'ID': 'Unique Code',
        'UserName': 'Head Department Entity',
        'Password': 'Secret Code',
        'userRole': 'Principle',
        'header': true
      },
      {
        'ID': 4,
        'UserName': 'User 4',
        'Password': 'Password4',
        'userRole': 'Principle'
      },
      {
        'ID': 8,
        'UserName': 'User 8',
        'Password': 'Password',
        'userRole': 'Principle'
      },
      {
        'ID': 'Unique Code',
        'UserName': 'Name Of Department Entity',
        'Password': 'Secret Code',
        'userRole': 'Role',
        'header': true
      },
      {
        'ID': 1,
        'UserName': 'User 1',
        'Password': 'Password1',
        'userRole': 'Student'
      },
      {
        'ID': 2,
        'UserName': 'User 2',
        'Password': 'Password2',
        'userRole': 'Teacher'
      },
      {
        'ID': 3,
        'UserName': 'User 3',
        'Password': 'Password3',
        'userRole': 'Student'
      },
      {
        'ID': 5,
        'UserName': 'User 5',
        'Password': 'Password5',
        'userRole': 'Student'
      },
      {
        'ID': 6,
        'UserName': 'User 6',
        'Password': 'Password6',
        'userRole': 'Student'
      },
      {
        'ID': 7,
        'UserName': 'User 7',
        'Password': 'Password7',
        'userRole': 'Student'
      },
      {
        'ID': 9,
        'UserName': 'User 9',
        'Password': 'Password9',
        'userRole': 'Teacher'
      },
      {
        'ID': 10,
        'UserName': 'User 10',
        'Password': 'Password10',
        'userRole': 'Teacher'
      }
    ];
    const headers = [{
      value: 'ID',
      width: 20
    }, {
      value: 'UserName',
      width: 40
    }, {
      value: 'Password',
      width: 30
    }, {
      value: 'userRole',
      width: 20
    }];
    this.excelService.exportAsExcelFile(this.data, 'abc', headers);
  }
}
