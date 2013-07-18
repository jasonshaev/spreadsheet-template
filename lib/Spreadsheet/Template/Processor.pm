package Spreadsheet::Template::Processor;
use Moose::Role;
# ABSTRACT: role for classes which preprocess a template file before rendering

requires 'process';

=head1 SYNOPSIS

  package MyProcessor;
  use Moose;

  with 'Spreadsheet::Template::Processor';

  sub process {
      # ...
  }

=head1 DESCRIPTION

This role should be consumed by any class which will be used as the
C<processor_class> in a L<Spreadsheet::Template> instance.

=cut

=method process($contents, $vars)

This method is required to be implemented by any classes which consume this
role. It should take the contents of the template and return a JSON file as
described in L<Spreadsheet::Template>. This typically just means running it
through a template engine of some kind.

=cut

no Moose::Role;

1;
