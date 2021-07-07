requires 'perl', '>= 5.020';

requires 'App::Cmd', '0.333';
requires 'Spreadsheet::ParseExcel', '0.65';
requires 'Excel::Writer::XLSX' , '1.09';

on test => sub {
    requires 'Test::More', '0.96';
};
