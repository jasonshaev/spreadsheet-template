package Spreadsheet::Template::Generator::Parser::XLS;
use Moose;

use Spreadsheet::ParseExcel;

with 'Spreadsheet::Template::Generator::Parser::Excel';

sub make_excel {
    my $self = shift;
    my ($filename) = @_;

    my $parser = Spreadsheet::ParseExcel->new;
    my $excel = $parser->parse($filename);
    die $parser->error unless $excel;

    return $excel;
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
