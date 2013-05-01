package Spreadsheet::Template::Generator::Parser::XLSX;
use Moose;

use Spreadsheet::XLSX;
use XML::Entities;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub make_excel {
    my $self = shift;
    my ($filename) = @_;
    return Spreadsheet::XLSX->new($filename);
}

sub _filter_cell_contents {
    my $self = shift;
    my ($contents) = @_;
    # XXX this decode call really feels like a bug in Spreadsheet::XLSX
    return XML::Entities::decode('all', $contents);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
