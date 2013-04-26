package Spreadsheet::Template;
use Moose;

use Class::Load 'load_class';
use JSON;

has processor_class => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Spreadsheet::Template::Processor::Xslate',
);

has writer_class => (
    is      => 'ro',
    isa     => 'Str',
    default => 'Spreadsheet::Template::Writer::XLSX',
);

has processor => (
    is      => 'ro',
    does    => 'Spreadsheet::Template::Processor',
    handles => 'Spreadsheet::Template::Processor',
    lazy    => 1,
    default => sub {
        my $self = shift;
        my $class = $self->processor_class;
        load_class($class);
        return $class->new;
    },
);

has writer => (
    is      => 'ro',
    does    => 'Spreadsheet::Template::Writer',
    handles => 'Spreadsheet::Template::Writer',
    lazy    => 1,
    default => sub {
        my $self = shift;
        my $class = $self->writer_class;
        load_class($class);
        return $class->new;
    },
);

sub render {
    my $self = shift;
    my ($template, $vars) = @_;
    my $contents = $self->process($template, $vars);
    # not decode_json, since we expect that we are already being handed a
    # character string (decode_json also decodes utf8)
    my $data = from_json($contents);
    return $self->write($data);
}

__PACKAGE__->meta->make_immutable;
no Moose;

1;
