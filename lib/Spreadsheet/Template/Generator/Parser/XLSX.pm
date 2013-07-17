package Spreadsheet::Template::Generator::Parser::XLSX;
use Moose;
# ABSTRACT: parser for XLSX files

use Spreadsheet::ParseXLSX;

with 'Spreadsheet::Template::Generator::Parser::Excel';

=head1 SYNOPSIS

  my $generator = Spreadsheet::Template::Generator->new(
      parser_class => 'Spreadsheet::Template::Generator::Parser',
  );

=head1 DESCRIPTION

This is an implementation of L<Spreadsheet::Template::Generator::Parser> for
XLSX files. It uses L<Spreadsheet::ParseXLSX> to do the parsing.

=cut

sub _create_workbook {
    my $self = shift;
    my ($filename) = @_;

    my $parser = Spreadsheet::ParseXLSX->new($filename);
    return $parser->parse($filename);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
