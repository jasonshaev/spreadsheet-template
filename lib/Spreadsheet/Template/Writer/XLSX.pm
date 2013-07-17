package Spreadsheet::Template::Writer::XLSX;
use Moose;
# ABSTRACT: generate XLSX files from templates

with 'Spreadsheet::Template::Writer::Excel';

sub excel_class { 'Excel::Writer::XLSX' }

__PACKAGE__->meta->make_immutable;
no Moose;

1;
