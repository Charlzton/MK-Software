import { PatientRecord, DVLARecord, MedsRecord, Sheet2Record, ColumnDefinition } from '../types';

export const PATIENTS_COLUMNS: ColumnDefinition[] = [
  { key: 'date', label: 'Date', type: 'date' },
  { key: 'recordedBy', label: 'Recorded_By', type: 'text' },
  { key: 'phoneNo', label: 'Phone_No', type: 'text' },
  { key: 'examinedBy', label: 'Examined_By', type: 'text' },
  { key: 'name', label: 'Name', type: 'text', required: true },
  { key: 'dateOfBirth', label: 'Date_of_Birth', type: 'text' },
  { key: 'testType', label: 'Test_Type', type: 'text' },
  { key: 'testFee', label: 'Test_Fee', type: 'number' },
  { key: 'dropsNo', label: 'Drops_No', type: 'number' },
  { key: 'medication', label: 'Medication', type: 'number' },
  { key: 'total', label: 'Total', type: 'number' },
  { key: 'glassTotal', label: 'Glass Total', type: 'number' },
  { key: 'paid', label: 'Paid', type: 'text' },
  { key: 'glassesBal', label: 'Glasses BAL', type: 'number' },
  { key: 'reviewDate', label: 'Review Date', type: 'date' },
  { key: 'reviewTest', label: 'Review Test', type: 'text' },
];

export const DVLA_COLUMNS: ColumnDefinition[] = [
  { key: 'date', label: 'Date', type: 'date' },
  { key: 'recordedBy', label: 'Recorded_By (Col 2)', type: 'text' },
  { key: 'contact', label: 'Contact', type: 'text' },
  { key: 'dic', label: 'DIC', type: 'text' },
  { key: 'patientName', label: "Patient's Name", type: 'text', required: true },
  { key: 'newFee', label: 'New', type: 'number' },
  { key: 'renewalFee', label: 'Renewal', type: 'number' },
  { key: 'notes', label: 'Remarks / Notes', type: 'text' },
];

export const INITIAL_PATIENT_RECORDS: PatientRecord[] = [
  { id: 'MK-PAT-0001', date: '10-Jan-25', recordedBy: 'Irene', phoneNo: '0507408280', examinedBy: 'Dr. Bediako', name: 'Ibrahim Shuaib Liman-Toma', dateOfBirth: '1/20/2023', testType: 'Standard', testFee: 100, dropsNo: 2, medication: 100, total: 200, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0002', date: '10-Jan-25', recordedBy: 'irene', phoneNo: '0507408280', examinedBy: 'Dr. Bediako', name: 'Zinabu Sahabi', dateOfBirth: '3/15/1985', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0003', date: '10-Jan-25', recordedBy: 'Irene', phoneNo: '0268548506', examinedBy: 'Dr. Bediako', name: 'Christiana Asantewaa Sakyi', dateOfBirth: '7/13/2005', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0004', date: '13-Jan-25', recordedBy: 'Irene', phoneNo: '0240690079', examinedBy: 'Dr. Danso', name: 'Emmanuel Abrah Adjei', dateOfBirth: '3/17/1983', testType: 'Standard', testFee: 0, dropsNo: 0, medication: 0, total: 0, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0005', date: '13-Jan-25', recordedBy: 'Irene', phoneNo: '0597021869', examinedBy: 'Dr. Danso', name: 'Joyce Gbafa', dateOfBirth: '8/31/2004', testType: 'Standard', testFee: 100, dropsNo: 3, medication: 150, total: 250, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0006', date: '15-Jan-25', recordedBy: 'Irene', phoneNo: '0532574051', examinedBy: 'Dr. Danso', name: 'Lorda Asante', dateOfBirth: '1/12/2007', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0007', date: '17-Jan-25', recordedBy: 'Irene', phoneNo: '0257899165', examinedBy: 'Dr. Bediako', name: 'Ruth Akuffo', dateOfBirth: '5/31/1993', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0008', date: '20-Jan-25', recordedBy: 'Priscilla', phoneNo: '0536763300', examinedBy: 'Dr. Danso', name: 'Samuel Bismark Awuli', dateOfBirth: '5/25/1966', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0009', date: '20-Jan-25', recordedBy: 'Priscilla', phoneNo: '0246868576', examinedBy: 'Dr. Danso', name: 'Meryld Koui-Toglo', dateOfBirth: '4/3/1992', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0010', date: '22-Jan-25', recordedBy: 'Priscilla', phoneNo: '0243036974', examinedBy: 'Dr. Danso', name: 'Mavis Osae-Bosompim', dateOfBirth: '11/2/1987', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0011', date: '22-Jan-25', recordedBy: 'Priscilla', phoneNo: '0555133029', examinedBy: 'Dr. Danso', name: 'Rosemary Quarcoo', dateOfBirth: '10-Jul-59', testType: 'screening', testFee: 0, dropsNo: 1, medication: 50, total: 50, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0012', date: '22-Jan-25', recordedBy: 'Priscilla', phoneNo: '0263620639', examinedBy: 'Dr. Danso', name: 'Ibrahim Issaka', dateOfBirth: '5/23/1993', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0013', date: '27-Jan-25', recordedBy: 'Priscilla', phoneNo: '0597925880', examinedBy: 'Dr. Danso', name: 'Habiba Ali Nariam', dateOfBirth: '4/19/2000', testType: 'Standard', testFee: 100, dropsNo: 3, medication: 150, total: 250, glassTotal: '', paid: '', glassesBal: '', reviewDate: '2/3/2025', reviewTest: '' },
  { id: 'MK-PAT-0014', date: '29-Jan-25', recordedBy: 'Priscilla', phoneNo: '0553699094', examinedBy: 'Dr. Danso', name: 'Esther Amoah', dateOfBirth: '6/6/2005', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0015', date: '29-Jan-25', recordedBy: 'Priscilla', phoneNo: '0558710822', examinedBy: 'Dr. Danso', name: 'Charity Sekley', dateOfBirth: '11/1/1934', testType: 'Standard', testFee: 100, dropsNo: 4, medication: 200, total: 300, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0016', date: '29-Jan-25', recordedBy: 'Priscilla', phoneNo: '0558793962', examinedBy: 'Dr. Danso', name: 'Christabel Mensah', dateOfBirth: '2/19/1998', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0017', date: '31-Jan-25', recordedBy: 'Priscilla', phoneNo: '0273286973', examinedBy: 'Dr. Atta', name: 'Kwaku Oduro Asare', dateOfBirth: '11/13/1985', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0018', date: '31-Jan-25', recordedBy: 'Priscilla', phoneNo: '0257899165', examinedBy: 'Dr. Attah', name: 'Ruth Akuffo', dateOfBirth: '5/31/1993', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 0, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0019', date: '3-Feb-25', recordedBy: 'Irene', phoneNo: '0244428791', examinedBy: 'Dr. Danso', name: 'Gina Opoku Agyemang', dateOfBirth: '12/12/1980', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0020', date: '3-Feb-25', recordedBy: 'Irene', phoneNo: '0599929188', examinedBy: 'Dr.Danso', name: 'Genius Oduraa Annor', dateOfBirth: '6/20/2009', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0021', date: '3-Feb-25', recordedBy: 'Irene', phoneNo: '0553812465', examinedBy: 'Dr.Danso', name: 'Claudia Oduraa Annor', dateOfBirth: '10/1/2003', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0022', date: '5-Feb-25', recordedBy: 'Priscilla', phoneNo: '0594453679', examinedBy: 'Dr.Danso', name: 'Samuel Boateng Bonsu', dateOfBirth: '2/24/2004', testType: 'Standard', testFee: 100, dropsNo: 2, medication: 100, total: 200, glassTotal: '', paid: '', glassesBal: '', reviewDate: '2/10/2025', reviewTest: '' },
  { id: 'MK-PAT-0023', date: '5-Feb-25', recordedBy: 'Priscilla', phoneNo: '0536713381', examinedBy: 'Dr. Danso', name: 'Gifty Amaakaro', dateOfBirth: '8/17/2002', testType: 'Standard', testFee: 100, dropsNo: 3, medication: 150, total: 250, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0024', date: '6-Feb-25', recordedBy: 'Priscilla', phoneNo: '0531575515', examinedBy: 'Dr.Joe', name: 'Deborah Asante', dateOfBirth: '00-00-93', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '2/20/2025', reviewTest: '' },
  { id: 'MK-PAT-0025', date: '6-Feb-25', recordedBy: 'Priscilla', phoneNo: '0243441202', examinedBy: 'Dr.Joe', name: 'Henry Frimpong Boateng', dateOfBirth: '5/15/1999', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: '', paid: '', glassesBal: '', reviewDate: '2/20/2025', reviewTest: '' },
  { id: 'MK-PAT-0026', date: '7-Feb-25', recordedBy: 'Irene', phoneNo: '0246705413', examinedBy: 'Dr Bediako', name: 'Joyce Adwoa Serwaa', dateOfBirth: '5/20/1957', testType: 'Standard', testFee: 100, dropsNo: 2, medication: 100, total: 200, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0027', date: '10-Feb-25', recordedBy: 'Irene', phoneNo: '0244796366', examinedBy: 'Dr. Danso', name: 'Yaa Adjeiwaa Enyan', dateOfBirth: '1/7/2010', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0028', date: '11-Feb-25', recordedBy: 'Priscilla', phoneNo: '0536503907', examinedBy: 'Dr.Joe', name: 'Omane Agyekum Apau', dateOfBirth: '3/15/2006', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0029', date: '12-Feb-25', recordedBy: 'Priscilla', phoneNo: '0595223507', examinedBy: 'Dr.Danso', name: 'Martha Dovi', dateOfBirth: '8/12/2003', testType: 'Standard', testFee: 100, dropsNo: 2, medication: 100, total: 200, glassTotal: '', paid: '', glassesBal: '', reviewDate: '2/19/2025', reviewTest: '' },
  { id: 'MK-PAT-0030', date: '19-Feb-25', recordedBy: 'Priscilla', phoneNo: '0257365349', examinedBy: 'Dr.Danso', name: 'Priscilla Adjoba', dateOfBirth: '13-09-2001', testType: 'Standard', testFee: 100, dropsNo: 2, medication: 100, total: 200, glassTotal: '', paid: 'Paid', glassesBal: 500, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0031', date: '25-Feb-25', recordedBy: 'Priscilla', phoneNo: '0244605421', examinedBy: 'Dr.Joe', name: 'Humpfrey Q Nunoo', dateOfBirth: '4/20/1939', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: 1150, paid: 'Paid', glassesBal: '', reviewDate: '3/11/2025', reviewTest: '' },
  { id: 'MK-PAT-0032', date: '3-Mar-25', recordedBy: 'Priscilla', phoneNo: '0552814814', examinedBy: 'Dr.Danso', name: 'Ama Attah', dateOfBirth: '9/14/1964', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: 1300, paid: 'Paid', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0033', date: '5-Mar-25', recordedBy: 'Priscilla', phoneNo: '0244510219', examinedBy: 'Dr. Danso', name: 'Betty Amoani', dateOfBirth: '5/14/1982', testType: 'Standard', testFee: '', dropsNo: '', medication: '', total: '', glassTotal: 1670, paid: 'Paid', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0034', date: '10-Mar-25', recordedBy: 'Priscilla', phoneNo: '0555496799', examinedBy: 'Dr.Danso', name: 'Katriel Eyison', dateOfBirth: '10/27/2015', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: 1200, paid: 'Paid', glassesBal: '', reviewDate: '3/24/2025', reviewTest: '' },
  { id: 'MK-PAT-0035', date: '10-Mar-25', recordedBy: 'Priscilla', phoneNo: '0242629590', examinedBy: 'Dr.Danso', name: 'William Kwame Zotorglo', dateOfBirth: '9/19/1943', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 700, paid: 'Paid', glassesBal: 350, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0036', date: '10-Mar-25', recordedBy: 'Priscilla', phoneNo: '0244384555', examinedBy: 'Dr. Danso', name: 'Prince Darko', dateOfBirth: '8/2/1972', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 1240, paid: 'Paid', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0037', date: '12-Mar-25', recordedBy: 'Priscilla', phoneNo: '0240881810', examinedBy: 'Dr.Danso', name: 'Vida Gbanchie', dateOfBirth: '4/4/1986', testType: 'Standard', testFee: 100, dropsNo: 2, medication: 100, total: 200, glassTotal: 600, paid: 'paid', glassesBal: 340, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0038', date: '19-Mar-25', recordedBy: 'Irene', phoneNo: '0553594904', examinedBy: 'Dr. Danso', name: 'Cynthia Bedzrah', dateOfBirth: '9/3/1998', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: 850, paid: 1000, glassesBal: 650, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0039', date: '19-Mar-25', recordedBy: 'Irene', phoneNo: '0270040976', examinedBy: 'Dr. Danso', name: 'Sheila Parku', dateOfBirth: '10/1/1989', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 1000, paid: 1000, glassesBal: 420, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0040', date: '24-Mar-25', recordedBy: 'Irene', phoneNo: '0244214232', examinedBy: 'Dr. Danso', name: 'Dorothea Halm Otoo', dateOfBirth: '3/31/1955', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 1550, paid: 1100, glassesBal: 550, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0041', date: '4-Apr-25', recordedBy: 'Irene', phoneNo: '0538978006', examinedBy: 'Dr. Levi', name: 'Jeffery Asiedu', dateOfBirth: '19-02-2006', testType: 'Standard', testFee: 250, dropsNo: 0, medication: 0, total: 250, glassTotal: 1350, paid: 1350, glassesBal: 0, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0042', date: '8-Apr-25', recordedBy: 'Irene', phoneNo: '0592000841', examinedBy: 'Dr. Joe', name: 'Maame Akua Danquah Danso', dateOfBirth: '25-Jun-97', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 1220, paid: 970, glassesBal: 350, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0043', date: '7-May-25', recordedBy: 'Irene', phoneNo: '0243290064', examinedBy: 'Dr Danso', name: 'Felicia Quainoo', dateOfBirth: '01-11-2008', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 1850, paid: 1850, glassesBal: 0, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0044', date: '12-May-25', recordedBy: 'Irene', phoneNo: '0247684512', examinedBy: 'Dr Danso', name: 'Evelyn Seyram Avorgbedor', dateOfBirth: '18-09-1989', testType: 'Standard', testFee: 100, dropsNo: 1, medication: 50, total: 150, glassTotal: 1300, paid: 650, glassesBal: 650, reviewDate: '26-05-2025', reviewTest: '' },
  { id: 'MK-PAT-0045', date: '21-May-25', recordedBy: 'Irene', phoneNo: '0554752112', examinedBy: 'Dr. Danso', name: 'Emelia Opoku', dateOfBirth: '28-12-1973', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 1350, paid: 900, glassesBal: 450, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0046', date: '23-May-25', recordedBy: 'Irene', phoneNo: '0244423077', examinedBy: 'Dr. Bediako', name: 'Kelly Osei', dateOfBirth: '01-01-2015', testType: 'Standard', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 1500, paid: 1000, glassesBal: 500, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0047', date: '21-Jul-25', recordedBy: 'Priscilla', phoneNo: '0240115050', examinedBy: 'Dr.Danso', name: 'Salamatu Mustapha', dateOfBirth: '4-8-1991', testType: 'Comprehensive', testFee: 350, dropsNo: 0, medication: 0, total: 350, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0048', date: '21-Jul-25', recordedBy: 'Priscilla', phoneNo: '0546150393', examinedBy: 'Dr. Danso', name: 'Akua Frema Asante-Agyei', dateOfBirth: '6-9-1978', testType: 'Standard', testFee: 150, dropsNo: 0, medication: 0, total: 150, glassTotal: 1870, paid: 1400, glassesBal: 470, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0049', date: '28-Jul-25', recordedBy: 'Priscilla', phoneNo: '0268691959', examinedBy: 'Dr Danso', name: 'Yaw Afriyie Owusu-Akyaw', dateOfBirth: '13-5-1993', testType: 'I . O . P', testFee: 100, dropsNo: 0, medication: 0, total: 100, glassTotal: 0, paid: 0, glassesBal: 0, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0050', date: '29-Jul-25', recordedBy: 'Priscilla', phoneNo: '0268691959', examinedBy: 'Dr. Lucy', name: 'Yaw Afriyie Owusu-Akyaw', dateOfBirth: '13-5-1993', testType: 'VfT', testFee: 150, dropsNo: 0, medication: 0, total: 150, glassTotal: 0, paid: 0, glassesBal: 0, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0051', date: '24-Sep-25', recordedBy: 'Irene', phoneNo: '0246333833', examinedBy: 'Dr. Danso', name: 'Nana Sam .K. Gyasi', dateOfBirth: '15-03-1962', testType: 'Standard', testFee: 230, dropsNo: 0, medication: 0, total: 230, glassTotal: 3580, paid: 1800, glassesBal: 1580, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0052', date: '27-Oct-25', recordedBy: 'Irene', phoneNo: '0201404558', examinedBy: 'Dr. Danso', name: 'Nana Abrah Adjei', dateOfBirth: '25-05-1955', testType: 'Standard', testFee: 150, dropsNo: 0, medication: 0, total: 150, glassTotal: 3750, paid: 3750, glassesBal: 0, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0053', date: '14-Jan-26', recordedBy: 'Irene', phoneNo: '0247894443', examinedBy: 'Dr Danso', name: 'Janet Owoo', dateOfBirth: '17-05-1962', testType: 'Standard', testFee: 150, dropsNo: 1, medication: 50, total: 200, glassTotal: 1700, paid: 1100, glassesBal: 600, reviewDate: '246950148', reviewTest: 'Mr Desmond' },
  { id: 'MK-PAT-0054', date: '4-Feb-26', recordedBy: 'Irene', phoneNo: '0243675332', examinedBy: 'Dr Danso', name: 'Rita Barketey', dateOfBirth: '18-09-1982', testType: 'Standard', testFee: 150, dropsNo: 0, medication: 0, total: 0, glassTotal: 2400, paid: 2400, glassesBal: 0, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0055', date: '27-Mar-26', recordedBy: 'Charles', phoneNo: '0205107300', examinedBy: 'Dr. Herbert', name: 'Clare Emefa Dabi', dateOfBirth: '29-10-2001', testType: 'Standard', testFee: 150, dropsNo: '', medication: 50, total: '', glassTotal: 1000, paid: 'Paid', glassesBal: 0, reviewDate: '27-05-2026', reviewTest: 'IOP' },
  { id: 'MK-PAT-0056', date: '27-Mar-26', recordedBy: 'Charles', phoneNo: '0257365349', examinedBy: 'Dr. Herbert', name: 'Priscilla Adjoba Cudjoe', dateOfBirth: '13-09-2001', testType: 'Comprehensive', testFee: 350, dropsNo: '', medication: 100, total: 450, glassTotal: '', paid: 'Paid', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0057', date: '13-Apr-26', recordedBy: 'Priscilla', phoneNo: '0554007825', examinedBy: 'Dr Danso', name: 'Dominic Boakye', dateOfBirth: '2-Apr-09', testType: 'Standard', testFee: 100, dropsNo: '', medication: '', total: 150, glassTotal: 1570, paid: 1570, glassesBal: 0, reviewDate: '17-Jun-26', reviewTest: 'VFT' },
  { id: 'MK-PAT-0058', date: '1-Jul-26', recordedBy: 'Charles', phoneNo: '0550684684', examinedBy: 'Dr Boatemaa', name: 'Abdul Latif', dateOfBirth: '1-Mar-72', testType: 'Standard', testFee: 150, dropsNo: '', medication: '', total: '', glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0059', date: '10-Jul-26', recordedBy: 'Charles', phoneNo: '0243674168', examinedBy: 'Dr Danso', name: 'Lydia Amponsah', dateOfBirth: '25-Mar-81', testType: 'Standard', testFee: '', dropsNo: '', medication: '', total: '', glassTotal: 1700, paid: 700, glassesBal: 1000, reviewDate: '', reviewTest: '' },
  { id: 'MK-PAT-0060', date: '13-Oct-26', recordedBy: 'Charles', phoneNo: '0507466181', examinedBy: 'Dr Danso', name: 'Joshua Adjei', dateOfBirth: '23-May-03', testType: 'Standard', testFee: 150, dropsNo: 0, medication: 0, total: 150, glassTotal: '', paid: '', glassesBal: '', reviewDate: '', reviewTest: '' }
];

export const INITIAL_DVLA_RECORDS: DVLARecord[] = [
  { id: 'MK-DVLA-0001', date: '2-Apr-25', recordedBy: 'Priscilla Asare', contact: '0543861680', dic: 'Dr Danso', patientName: 'Azangbaatanaab A. Gideon', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0002', date: '4-Apr-25', recordedBy: 'Priscilla', contact: '0597156896', dic: 'Dr Levi', patientName: 'Eric Abam', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0003', date: '24-Apr-25', recordedBy: 'Priscilla', contact: '0242965190', dic: 'Dr Joe', patientName: 'Nana Ama Konadu Asubonteng', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0004', date: '2-May-25', recordedBy: 'Priscilla', contact: '0279671978', dic: 'Dr Attah', patientName: 'Jephter Odame', newFee: 0, renewalFee: 50, notes: '' },
  { id: 'MK-DVLA-0005', date: '5-May-25', recordedBy: 'irene', contact: '0244480539', dic: 'Dr Danso', patientName: 'Linda Asiamah', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0006', date: '12-May-25', recordedBy: 'Priscilla', contact: '0244829215', dic: 'Dr Danso', patientName: 'Asante Benjamin', newFee: 0, renewalFee: 50, notes: '' },
  { id: 'MK-DVLA-0007', date: '14-May-25', recordedBy: 'irene', contact: '0557382500', dic: 'Dr Danso', patientName: 'Vincent Epanty Mofa', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0008', date: '16-May-25', recordedBy: 'irene', contact: '0279101610', dic: 'Dr Bediako', patientName: 'Mustapha Musah', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0009', date: '23-May-25', recordedBy: 'Priscilla Asare', contact: '0597150772', dic: 'Dr Bediako', patientName: 'Hafiz Inusah', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0010', date: '2-Jun-25', recordedBy: 'Priscilla Asare', contact: '0546324882', dic: 'Dr Danso', patientName: 'Mensah John', newFee: 70, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0011', date: '13-Jun-25', recordedBy: 'irene', contact: '0554663453', dic: 'Dr. Bediako', patientName: 'John Osae', newFee: 100, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0012', date: '20-Jun-25', recordedBy: 'Priscilla', contact: '0208551580', dic: 'Dr. Bediako', patientName: 'Mensah Samuel Kwadwo', newFee: 0, renewalFee: 80, notes: '' },
  { id: 'MK-DVLA-0013', date: '23-Jun-25', recordedBy: 'Priscilla Asare', contact: '0558206611', dic: 'Dr Danso', patientName: 'Iddrisu Yussif', newFee: 100, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0014', date: '8-Jul-25', recordedBy: 'Priscilla', contact: '0244324489', dic: 'Dr. Lucy', patientName: 'Korley Ben Nii', newFee: 0, renewalFee: 80, notes: '' },
  { id: 'MK-DVLA-0015', date: '18-Jul-25', recordedBy: 'Priscilla', contact: '0202110091', dic: 'Dr. Bediako', patientName: 'Sarah Akuamoah Boateng', newFee: 100, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0016', date: '30-Jul-25', recordedBy: 'irene', contact: '0244714209', dic: 'Dr. Lucy', patientName: 'Joseph Osei Agyepong', newFee: 80, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0017', date: '4-Aug-25', recordedBy: 'irene', contact: '0501596285', dic: 'Dr. Danso', patientName: 'Akaba Malcolm Jewitt', newFee: 0, renewalFee: 80, notes: '' },
  { id: 'MK-DVLA-0018', date: '25-Aug-25', recordedBy: 'irene', contact: '0242011921', dic: 'Dr. Danso', patientName: 'Issaka Alafa', newFee: 100, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0019', date: '15-Apr-26', recordedBy: 'Priscilla', contact: '0249599164', dic: 'Dr Danso', patientName: 'Anaba Mavis', newFee: 100, renewalFee: 0, notes: 'RW' },
  { id: 'MK-DVLA-0020', date: '20-Apr-26', recordedBy: 'Charles', contact: '0576497750', dic: 'Dr Danso', patientName: 'Andrew awuni', newFee: 100, renewalFee: 0, notes: 'RW' },
  { id: 'MK-DVLA-0021', date: '27-Apr-26', recordedBy: 'Priscilla', contact: '0502945219', dic: 'Dr Danso', patientName: 'Tsagli Paul-Dag K. M.', newFee: 100, renewalFee: 0, notes: 'RW' },
  { id: 'MK-DVLA-0022', date: '3-Jun-26', recordedBy: 'PrIscilla', contact: '0530834803', dic: 'Dr Danso', patientName: 'Akplene Senoo Atsu', newFee: 100, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0023', date: '8-Jun-26', recordedBy: 'Charles', contact: '0208224046', dic: 'Dr Danso', patientName: 'Tudzi Justice Edem Kwame', newFee: 0, renewalFee: 80, notes: '' },
  { id: 'MK-DVLA-0024', date: '22-Jul-26', recordedBy: 'priscilla', contact: '0592278147', dic: 'Dr Matthew', patientName: 'Amadu Matina', newFee: 100, renewalFee: 0, notes: '' },
  { id: 'MK-DVLA-0025', date: '14-Sep-26', recordedBy: 'Priscilla', contact: '0599092688', dic: 'Dr Joshua', patientName: 'Amankwa Stephen Nana Kofi', newFee: 100, renewalFee: 0, notes: '' }
];

export const INITIAL_MEDS_RECORDS: MedsRecord[] = [
  { id: 'MK-MED-0001', medicineName: 'NACROM', val1: 300, val2: 10, val3: 1, val4: 2 },
  { id: 'MK-MED-0002', medicineName: 'OCULOSE', val1: 100, val2: 12, val3: 5, val4: 7 },
  { id: 'MK-MED-0003', medicineName: 'EYECROM', val1: 50, val2: 5, val3: 2, val4: 1 },
  { id: 'MK-MED-0004', medicineName: 'CIVOX', val1: 5, val2: 5, val3: 0, val4: 2 },
  { id: 'MK-MED-0005', medicineName: 'TIMOLOL', val1: 3, val2: 0, val3: 0, val4: 0 },
];

export const INITIAL_SHEET2_RECORDS: Sheet2Record[] = [
  { id: 'MK-S2-0001', col1: 'Esther Owusu Mensah', col2: '', col3: '', contact: '0536395422', col5: '', date: '3/11/2025' },
  { id: 'MK-S2-0002', col1: 'Georgina Mensah', col2: '', col3: '', contact: '0246383896', col5: '', date: '5/9/2025' },
  { id: 'MK-S2-0003', col1: 'Paul Afful', col2: '', col3: '', contact: '0244067287', col5: '', date: '5/27/1967' },
];

export const KNOWN_DOCTORS = [
  'Dr. Danso',
  'Dr. Bediako',
  'Dr. Lucy',
  'Dr. Joe',
  'Dr. Attah',
  'Dr. Levi',
  'Dr. Matthew',
  'Dr. Joshua',
  'Dr. Herbert',
  'Dr. Boatemaa',
  'Dr. Emily',
  'Dr. Deelyn',
  'Dr. Abena'
];

export const KNOWN_TEST_TYPES = [
  'Standard',
  'Comprehensive',
  'Review',
  'I . O . P',
  'VfT',
  'screening'
];

export const KNOWN_STAFF = [
  'Priscilla',
  'Irene',
  'Charles',
  'Priscilla Asare'
];
