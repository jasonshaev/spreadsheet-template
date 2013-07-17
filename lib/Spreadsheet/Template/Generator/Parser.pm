package Spreadsheet::Template::Generator::Parser;
use Moose::Role;
# ABSTRACT: role for classes which parse an existing spreadsheet

requires 'parse';

=head1 SYNOPSIS

  package MyParser;
  use Moose;

  with 'Spreadsheet::Template::Generator::Parser';

  sub parse {
      # ...
  }

=head1 DESCRIPTION

This role should be consumed by any class which will be used as the
C<parser_class> in a L<Spreadsheet::Template::Generator> instance.

=cut

=method parse($filename) (required)

This method should parse the spreadsheet specified by C<$filename> and return
the intermediate data structure containing all of the data in that spreadsheet.
The intermediate data format is documented in L<Spreadsheet::Template>.

=cut

no Moose::Role;

1;
