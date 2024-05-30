import { Injectable } from '@nestjs/common';
import * as moment from 'moment';
import { camalize, getDaysInMonth } from './utils';
import { google } from 'googleapis';
import 'moment/locale/vi';
moment.locale('vi');
import 'moment-timezone';

@Injectable()
export class AppService {
  getHello(): string {
    return 'Hello World!';
  }

  processData({ month, year }, rawData, idMapping) {
    const employees = idMapping.map((row) => {
      const records = rawData
        .filter((e) => e[0] == row[2])
        .map((e) => new Date(e[1]));
      console.log({ records });

      return {
        name: row[1],
        emp_id: row[2],
        records: records,
      };
    });
    const daysInMonth = getDaysInMonth(month, year);
    employees.forEach((employee) => {
      console.log({ employee });

      employee['perDays'] = [];
      daysInMonth.forEach((day: Date) => {
        const rowForDate = {
          day: moment(day).tz('Asia/Ho_Chi_Minh').format('YYYY-MM-DD'),
          dayOfWeek: camalize(
            moment(day).tz('Asia/Ho_Chi_Minh').format('dddd'),
          ),
          standard: '08:00',
          isHoliday: [0, 6].includes(moment(day).day()),
        };
        const recordsOfDay = employee.records
          .map((e) => e.valueOf())
          .filter(
            (e) => e > day.valueOf() && e < day.valueOf() + 3600 * 24 * 1000,
          );
        console.log({ recordsOfDay });

        if (recordsOfDay.length == 0) {
          // absent
          rowForDate['from'] = '';
          rowForDate['to'] = '';
        } else if (recordsOfDay.length == 1) {
          // forgot to check in/out
          rowForDate['overTime'] = '00:00:00';
          rowForDate['overTimeMs'] = 0;
          if (recordsOfDay[0] < day.valueOf() + 3600 * 12.5 * 1000) {
            // checked in, forgot to checkout
            rowForDate['from'] = moment(new Date(recordsOfDay[0])).format(
              'HH:mm:ss',
            );
            rowForDate['to'] = moment(
              new Date(day.valueOf() + 3600 * 17 * 1000),
            ).format('HH:mm:ss');
          } else {
            // checked out, forgot to checkin
            rowForDate['from'] = moment(
              new Date(day.valueOf() + 3600 * 8 * 1000),
            ).format('HH:mm:ss');
            rowForDate['to'] = moment(new Date(recordsOfDay[0])).format(
              'HH:mm:ss',
            );
          }
        } else {
          const ealiest = new Date(Math.min(...recordsOfDay));
          const latest = new Date(Math.max(...recordsOfDay));
          rowForDate['from'] = moment(ealiest).format('HH:mm:ss');
          rowForDate['to'] = moment(latest).format('HH:mm:ss');
        }
        employee['perDays'].push(rowForDate);
      });
    });
    return employees;
  }

  async createSheet({ month, year }) {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const key = require(process.cwd() + '/google-services-key.json');
    // Get credentials and build service
    // TODO (developer) - Use appropriate auth mechanism for your app
    const authClient = new google.auth.JWT(
      key['client_email'],
      null,
      key['private_key'],
      'https://www.googleapis.com/auth/drive',
    );
    const service = google.drive({ version: 'v3', auth: authClient });
    try {
      console.log('create file------------------------');
      const file = await service.files.create({
        resource: {
          name: `${('0' + month).slice(-2)} - ${year}`,
          parents: [`${process.env.FOLDER_DRIVE_ID}`],
          mimeType: 'application/vnd.google-apps.spreadsheet',
        },
      } as any);
      console.log('File:', file);
      console.log('File Id:', file.data.id);
      return file.data.id;
    } catch (err) {
      // TODO(developer) - Handle error
      throw err;
    }
  }

  async generateSheet({ month, year }, employees) {
    // eslint-disable-next-line @typescript-eslint/no-var-requires
    const key = require(process.cwd() + '/google-services-key.json');
    const sheetId = await this.createSheet({ month, year });
    const authClient = new google.auth.JWT(
      key['client_email'],
      null,
      key['private_key'],
      'https://www.googleapis.com/auth/spreadsheets',
    );
    const sheets = google.sheets({ version: 'v4', auth: authClient });
    const spreadsheetData = (
      await sheets.spreadsheets.get({
        spreadsheetId: sheetId,
        auth: authClient,
      })
    ).data;
    const sheetNames = spreadsheetData.sheets.map((e) => e.properties.title);
    for (const employee of employees) {
      employee['sheetName'] = `${employee.name} (${employee.emp_id})`;
    }
    const sheetsToCreate = employees
      .filter((e) => !sheetNames.includes(e['sheetName']))
      .map((e) => e['sheetName']);
    if (sheetsToCreate.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: {
          requests: sheetsToCreate.map((e) => ({
            addSheet: {
              properties: {
                title: e,
              },
            },
          })),
        },
      });
    }
    const sheetsData = (
      await sheets.spreadsheets.get({
        spreadsheetId: sheetId,
        auth: authClient,
      })
    ).data;
    console.log(sheetsData);
    const sheetToDelete = [];
    for (const sheet of sheetsData.sheets) {
      const shouldDelete = !employees
        .map((e) => e['sheetName'])
        .includes(sheet.properties.title);
      if (shouldDelete) {
        sheetToDelete.push(sheet.properties.sheetId);
      } else {
        const emp = employees.find(
          (e) => e['sheetName'] == sheet.properties.title,
        );
        emp['sheetId'] = sheet.properties.sheetId;
      }
    }
    if (sheetToDelete.length > 0) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: sheetId,
        requestBody: {
          requests: sheetToDelete.map((e) => ({
            deleteSheet: {
              sheetId: e,
            },
          })),
        },
      });
    }
    const batch = [];
    const styleBatch = [];
    for (const employee of employees) {
      batch.push({
        range: `${employee['sheetName']}!A1:A3`,
        values: [
          ['CHẤM CÔNG'],
          [`Mã số nhân viên (${employee.emp_id}): ${employee.name}`],
          [`Bảng chấm công tháng ${month}-${year}`],
        ],
      });
      batch.push({
        range: `${employee['sheetName']}!A4:L4`,
        values: [
          [
            'Ngày',
            'Thứ',
            'Check in',
            'Check out',
            'Thời gian làm việc',
            'Thời gian nghỉ trưa',
            'Thời gian làm việc tiêu chuẩn',
            'Làm quá giờ',
            'Làm không đủ giờ',
            'Ngày công',
            'Phép được duyệt',
            'Ghi chú',
          ],
        ],
      });
      let rowIndex = 5;
      for (const dayData of employee['perDays']) {
        if (!(dayData['isHoliday'] && !dayData['from'] && !dayData['to'])) {
          batch.push({
            range: `${employee['sheetName']}!A${rowIndex}:L${rowIndex}`,
            values: [
              [
                dayData['day'],
                dayData['dayOfWeek'],
                dayData['from'] ? `=TIMEVALUE("${dayData['from']}")` : '',
                dayData['from'] ? `=TIMEVALUE("${dayData['to']}")` : '',
                `=D${rowIndex}-C${rowIndex}-F${rowIndex}`, //dayData[`standard`],
                `=IF(AND(C${rowIndex}<TIMEVALUE("12:00:00");D${rowIndex}>TIMEVALUE("13:00:00"));TIMEVALUE("01:00:00");TIMEVALUE("00:00:00"))`, //dayData[`overTime`],
                dayData['isHoliday']
                  ? `=TIMEVALUE("00:00:00")`
                  : `=TIMEVALUE("08:00:00")`, // dayData[`missingTime`],
                `=IF(E${rowIndex}-G${rowIndex}>0;E${rowIndex}-G${rowIndex};TIMEVALUE("00:00:00"))`,
                `=IF(G${rowIndex}-E${rowIndex}>0;G${rowIndex}-E${rowIndex};TIMEVALUE("00:00:00"))`,
                `=IF(E${rowIndex}>TIMEVALUE("07:00:00");1;IF(E${rowIndex}>TIMEVALUE("04:00:00");0.5;0))`,
                '',
                dayData['note'],
              ],
            ],
          });
        } else {
          batch.push({
            range: `${employee['sheetName']}!A${rowIndex}:L${rowIndex}`,
            values: [
              [
                dayData['day'],
                dayData['dayOfWeek'],
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                dayData['note'],
              ],
            ],
          });
        }
        if (dayData['isHoliday']) {
          styleBatch.push({
            repeatCell: {
              range: {
                sheetId: employee['sheetId'],
                startRowIndex: rowIndex - 1,
                endRowIndex: rowIndex,
                startColumnIndex: 0,
                endColumnIndex: 12,
              },
              cell: {
                userEnteredFormat: {
                  backgroundColor: {
                    red: 1,
                    green: 0.88,
                    blue: 0.8,
                  },
                },
              },
              fields: 'userEnteredFormat(backgroundColor)',
            },
          });
        }
        rowIndex += 1;
      }
      batch.push({
        range: `${employee['sheetName']}!G${rowIndex + 1}:K${rowIndex + 2}`,
        values: [
          [
            'Tổng',
            `=SUM(H5:H${rowIndex})`,
            `=SUM(I5:I${rowIndex})`,
            `=SUM(J5:J${rowIndex})`,
            'Ngày công',
          ],
          [
            'Đánh giá',
            '=IF(H37>I37;"Làm thừa giờ"; "Làm thiếu giờ")',
            '=IF(H37>I37;H37-I37;I37-H37)',
            `=SUM(K5:K${rowIndex})`,
            'OT',
          ],
        ],
      });
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: rowIndex,
            endRowIndex: rowIndex + 2,
            startColumnIndex: 6,
            endColumnIndex: 11,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 1,
                green: 0.88,
                blue: 0.8,
              },
            },
          },
          fields: 'userEnteredFormat(backgroundColor)',
        },
      });
      styleBatch.push({
        updateBorders: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 3,
            endRowIndex: rowIndex - 1,
            startColumnIndex: 0,
            endColumnIndex: 12,
          },
          left: {
            style: 'SOLID_MEDIUM',
          },
          right: {
            style: 'SOLID_MEDIUM',
          },
          bottom: {
            style: 'SOLID_MEDIUM',
          },
          top: {
            style: 'SOLID_MEDIUM',
          },
          innerHorizontal: {
            style: 'SOLID_MEDIUM',
          },
          innerVertical: {
            style: 'SOLID_MEDIUM',
          },
        },
      });
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 0,
            startColumnIndex: 0,
          },
          cell: {
            userEnteredFormat: {
              textFormat: {
                fontFamily: 'Times New Roman',
              },
            },
          },
          fields: 'userEnteredFormat(textFormat(fontFamily))',
        },
      });
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 0,
            endRowIndex: 3,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          cell: {
            userEnteredFormat: {
              backgroundColor: {
                red: 1.0,
                green: 1.0,
                blue: 1.0,
              },
              horizontalAlignment: 'LEFT',
              textFormat: {
                foregroundColor: {
                  red: 0.26,
                  green: 0.52,
                  blue: 0.96,
                },
                fontSize: 14,
                bold: true,
                italic: true,
              },
            },
          },
          fields:
            'userEnteredFormat(backgroundColor,textFormat(foregroundColor,fontSize,bold,italic),horizontalAlignment)',
        },
      });
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 3,
            endRowIndex: 4,
            startColumnIndex: 0,
            endColumnIndex: 12,
          },
          cell: {
            userEnteredFormat: {
              textFormat: {
                bold: true,
              },
            },
          },
          fields: 'userEnteredFormat(textFormat(bold))',
        },
      });

      // date column
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 3,
            endRowIndex: rowIndex + 5,
            startColumnIndex: 0,
            endColumnIndex: 1,
          },
          cell: {
            userEnteredFormat: {
              horizontalAlignment: 'CENTER',
              textFormat: {
                fontSize: 12,
              },
              numberFormat: {
                type: 'DATE_TIME',
                pattern: 'DD/MM/YYYY',
              },
            },
          },
          fields:
            'userEnteredFormat(horizontalAlignment,textFormat(fontSize),numberFormat(pattern,type))',
        },
      });

      // day of week columns
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 3,
            endRowIndex: rowIndex + 5,
            startColumnIndex: 1,
            endColumnIndex: 2,
          },
          cell: {
            userEnteredFormat: {
              horizontalAlignment: 'CENTER',
              textFormat: {
                fontSize: 12,
              },
              numberFormat: {
                type: 'TEXT',
              },
            },
          },
          fields:
            'userEnteredFormat(horizontalAlignment,textFormat(fontSize),numberFormat(pattern,type))',
        },
      });

      // time columns
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 3,
            endRowIndex: rowIndex + 5,
            startColumnIndex: 2,
            endColumnIndex: 9,
          },
          cell: {
            userEnteredFormat: {
              horizontalAlignment: 'CENTER',
              textFormat: {
                fontSize: 12,
              },
              numberFormat: {
                type: 'DATE_TIME',
                pattern: 'hh:mm',
              },
            },
          },
          fields:
            'userEnteredFormat(horizontalAlignment,textFormat(fontSize),numberFormat(pattern,type))',
        },
      });
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 3,
            endRowIndex: rowIndex + 5,
            startColumnIndex: 9,
            endColumnIndex: 10,
          },
          cell: {
            userEnteredFormat: {
              horizontalAlignment: 'CENTER',
              textFormat: {
                fontSize: 12,
              },
              numberFormat: {
                type: 'NUMBER',
                pattern: '0.0',
              },
            },
          },
          fields:
            'userEnteredFormat(horizontalAlignment,textFormat(fontSize),numberFormat(pattern,type))',
        },
      });
      // format summary table
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: rowIndex + 1,
            endRowIndex: rowIndex + 2,
            startColumnIndex: 7,
            endColumnIndex: 9,
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'DATE_TIME',
                pattern: 'hh:mm',
              },
            },
          },
          fields: 'userEnteredFormat(numberFormat(pattern,type))',
        },
      });
      styleBatch.push({
        repeatCell: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: rowIndex + 2,
            endRowIndex: rowIndex + 3,
            startColumnIndex: 8,
            endColumnIndex: 9,
          },
          cell: {
            userEnteredFormat: {
              numberFormat: {
                type: 'DATE_TIME',
                pattern: 'hh:mm',
              },
            },
          },
          fields: 'userEnteredFormat(numberFormat(pattern,type))',
        },
      });

      styleBatch.push({
        mergeCells: {
          range: {
            sheetId: employee['sheetId'],
            startRowIndex: 0,
            endRowIndex: 3,
            startColumnIndex: 0,
            endColumnIndex: 4,
          },
          mergeType: 'MERGE_ROWS',
        },
      });
      styleBatch.push({
        autoResizeDimensions: {
          dimensions: {
            sheetId: employee['sheetId'],
            dimension: 'COLUMNS',
            startIndex: 1,
            endIndex: 12,
          },
        },
      });
    }
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: sheetId,
      auth: authClient,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: batch,
      },
    });
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: sheetId,
      requestBody: {
        requests: styleBatch,
      },
    });
    return `https://docs.google.com/spreadsheets/d/${sheetId}`;
  }
}
