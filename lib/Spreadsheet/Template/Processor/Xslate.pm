package Spreadsheet::Template::Processor::Xslate;
use Moose;

use Text::Xslate;

with 'Spreadsheet::Template::Processor';

has xslate => (
    is      => 'ro',
    isa     => 'Text::Xslate',
    lazy    => 1,
    default => sub { Text::Xslate->new(type => 'text') },
);

sub process {
    my $self = shift;
    my ($contents, $vars) = @_;
    return $self->xslate->render_string($contents, $vars);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
