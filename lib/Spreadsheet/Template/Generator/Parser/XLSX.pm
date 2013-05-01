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

# XXX this stuff all feels like working around bugs in Spreadsheet::XLSX -
# maybe look into that at some point
sub _filter_cell_contents {
    my $self = shift;
    my ($contents, $type) = @_;

    $contents = XML::Entities::decode('all', $contents);

    if ($type eq 'number') {
        $contents = 0+$contents;
    }

    return $contents;
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
