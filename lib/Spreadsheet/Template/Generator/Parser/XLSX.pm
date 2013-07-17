package Spreadsheet::Template::Generator::Parser::XLSX;
use Moose;
# ABSTRACT: parser for XLSX files

use Spreadsheet::ParseXLSX;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub _build_excel {
    my $self = shift;

    my $parser = Spreadsheet::ParseXLSX->new($self->filename);
    return $parser->parse($self->filename);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
