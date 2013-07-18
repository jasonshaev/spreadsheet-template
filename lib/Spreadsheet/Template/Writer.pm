package Spreadsheet::Template::Writer;
use Moose::Role;
# ABSTRACT: role for classes which write spreadsheet files from a template

requires 'write';

=head1 SYNOPSIS

  package MyWriter;
  use Moose;

  with 'Spreadsheet::Template::Writer';

  sub write {
      # ...
  }

=head1 DESCRIPTION

This role should be consumed by any class which will be used as the
C<writer_class> in a L<Spreadsheet::Template> instance.

=cut

=method write($data)

This method is required to be implemented by any classes which consume this
role. It should use the data in C<$data> (in the format described in
L<Spreadsheet::Template>) to create a new spreadsheet file containing that
data. It should return a string containing the binary contents of the
spreadsheet file.

=cut

no Moose::Role;

1;
