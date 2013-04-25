package Spreadsheet::Template::Writer::XLS;
use Moose;

with 'Spreadsheet::Template::Writer::Excel';

sub excel_class { 'Spreadsheet::WriteExcel' }

__PACKAGE__->meta->make_immutable;
no Moose;

1;
