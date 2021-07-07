package App::xls2xlsx;

use strict;
use warnings;

our $VERSION = '0.1.0';

use base qw( App::Cmd::Simple );

use Spreadsheet::ParseExcel;
use Excel::Writer::XLSX;

sub opt_spec {
  return (
    [ "clean|c", "clean up source file after successful conversion" ]
  );
}

sub execute {
  my ( $self, $opt, $args ) = @_;

  unless ( @{$args} ) {
    $self->usage_error("invalid file");
  }

  for my $filename ( @{$args} ) {
    unless ( -r -e $filename ) {
      _print_error( $filename . ' does not exist or has invalid permissions.' );
      next;
    }

    if ( _do_conversion($filename) ) {

      if ( $opt->{'clean'} ) {
        print 'Removing source file (' . $filename . ")\n";
        unless ( unlink $filename ) {
          print "failed to delete source file.\n";
        }
      }

      print 'Success: ' . $filename . "\n";
    }
  }

  return;

}

sub _do_conversion {
  my $filename = shift;

  my $xls_workbook  = _get_xls_workbook($filename);
  my $xlsx_workbook = _get_xlsx_workbook($filename);

  for my $xls_worksheet ( $xls_workbook->worksheets() ) {
    my $xlsx_worksheet = $xlsx_workbook->add_worksheet( $xls_worksheet->get_name() );

    my ( $row_min, $row_max ) = $xls_worksheet->row_range();
    my ( $col_min, $col_max ) = $xls_worksheet->col_range();

    for my $row ( $row_min .. $row_max ) {
      for my $col ( $col_min .. $col_max ) {
        my $xls_cell = $xls_worksheet->get_cell( $row, $col );
        next unless $xls_cell;
        my $format = _convert_xls_format( $xls_cell->get_format(), $xlsx_workbook );
        $xlsx_worksheet->write( $row, $col, $xls_cell->value(), $format );
      }
    }
  }

  $xlsx_workbook->close();

  return 1;
}

sub _convert_xls_format {
  my ( $xls_format, $xlsx_workbook ) = @_;

  unless ( ref($xls_format) eq 'Spreadsheet::ParseExcel::Format' ) {
    return;
  }

  my $format = $xlsx_workbook->add_format();

  $format->set_format_properties(
    font           => $xls_format->{'Font'}->{'Name'},
    size           => $xls_format->{'Font'}->{'Height'},
    color          => $xls_format->{'Font'}->{'Color'},
    bold           => $xls_format->{'Font'}->{'Bold'}   || 0,
    italic         => $xls_format->{'Font'}->{'Italic'} || 0,
    underline      => $xls_format->{'Font'}->{'UnderlineStyle'},
    font_strikeout => $xls_format->{'Font'}->{'Strikeout'} || 0,
    font_script    => $xls_format->{'Font'}->{'Super'}     || 0,
    align          => $xls_format->{'AlignH'}              || 0,
    valign         => $xls_format->{'AlignV'}              || 0,
    indent         => $xls_format->{'Indent'}              || 0,
    text_wrap      => $xls_format->{'Wrap'}                || 0,
    shrink         => $xls_format->{'Shrink'}              || 0,
    rotation       => $xls_format->{'Rotate'}              || 0,
    text_justlast  => $xls_format->{'JustLast'}            || 0,
    left           => @{ $xls_format->{'BdrStyle'} }[0],
    right          => @{ $xls_format->{'BdrStyle'} }[1],
    top            => @{ $xls_format->{'BdrStyle'} }[2],
    bottom         => @{ $xls_format->{'BdrStyle'} }[3],
    left_color     => @{ $xls_format->{'BdrColor'} }[0],
    right_color    => @{ $xls_format->{'BdrColor'} }[1],
    top_color      => @{ $xls_format->{'BdrColor'} }[2],
    bottom_color   => @{ $xls_format->{'BdrColor'} }[3],
    pattern        => @{ $xls_format->{'Fill'} }[0],
    fg_color       => @{ $xls_format->{'Fill'} }[1],
    bg_color       => @{ $xls_format->{'Fill'} }[2],
    locked         => $xls_format->{'Lock'}   || 0,
    hidden         => $xls_format->{'Hidden'} || 0,
  );

  return $format;
}

sub _get_xls_workbook {
  my $filename = shift;

  my $parser   = Spreadsheet::ParseExcel->new();
  my $workbook = $parser->parse($filename);

  unless ( ref($workbook) eq 'Spreadsheet::ParseExcel::Workbook' ) {
    if ( length( $parser->error() ) > 0 ) {
      die $parser->error() . ".\n";
    }
    else {
      die "XLS parser died from unknown error\n";
    }
  }

  return $workbook;
}

sub _get_xlsx_workbook {
  my $filename = shift;

  $filename =~ s/\.xls$/\.xlsx/;

  my $workbook = Excel::Writer::XLSX->new($filename);
  unless ( ref($workbook) eq 'Excel::Writer::XLSX' ) {
    die "Failure creating XLSX workbook\n";
  }

  return $workbook;
}

sub _print_error {
  my $message = shift;
  print "[ERROR] " . $message . "\n";
  return;
}

sub usage_error {
  my ( $self, $message ) = @_;

  print "Error: $message\n";
  print $self->usage_desc();

  exit 1;
}

sub usage_desc {
  return <<~ "EOF";
    usage: xls2xlsx [-?h] FILENAME [ FILENAME... ]
    EOF
}

1;

__END__

=encoding utf-8

=head1 NAME

App::xls2xlsx - Command line app for converting xls spreadsheets to xlsx spreasheets.

=head1 VERSION

version 0.1.0

=head1 AUTHOR

David Betz E<lt>hashref@gmail.comE<gt>

=head1 COPYRIGHT

This software is copyright (c) 2021 by David Betz

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
