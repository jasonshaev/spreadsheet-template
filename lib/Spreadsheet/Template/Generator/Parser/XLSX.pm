package Spreadsheet::Template::Generator::Parser::XLSX;
use Moose;
# ABSTRACT: parser for XLSX files

use Spreadsheet::ParseXLSX;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub _create_workbook {
    my $self = shift;
    my ($filename) = @_;

    my $parser = Spreadsheet::ParseXLSX->new($filename);
    return $parser->parse($filename);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
