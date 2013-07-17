package Spreadsheet::Template::Processor::Xslate;
use Moose;
# ABSTRACT: preprocess templates with Xslate

use Text::Xslate;

with 'Spreadsheet::Template::Processor';

has syntax => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Metakolon',
);

has xslate => (
    is      => 'ro',
    isa     => 'Text::Xslate',
    lazy    => 1,
    default => sub {
        my $self = shift;
        return Text::Xslate->new(
            type   => 'text',
            syntax => $self->syntax,
            module => ['Spreadsheet::Template::Helpers::Xslate'],
        );
    },
);

sub process {
    my $self = shift;
    my ($contents, $vars) = @_;
    return $self->xslate->render_string($contents, $vars);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
