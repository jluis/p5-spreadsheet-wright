package Spreadsheet::Wright::OpenDocumentXML;

use 5.010;
use strict;
use warnings;
no warnings qw( uninitialized numeric );

BEGIN {
	$Spreadsheet::Wright::OpenDocumentXML::VERSION   = '0.106';
	$Spreadsheet::Wright::OpenDocumentXML::AUTHORITY = 'cpan:TOBYINK';
}

use Carp;
use XML::LibXML;

use parent qw(Spreadsheet::Wright);
use constant {
	OFFICE_NS => "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
	STYLE_NS  => "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
	TEXT_NS   => "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
	TABLE_NS  => "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
	META_NS   => "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
	NUMBER_NS => "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
	CALCEXT_NS => "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
	};

sub new
{
	my ($class, %args) = @_;
	my $self = bless { 'options' => \%args }, $class;

	my $fh = $args{'fh'} // $args{'filehandle'};
	if ($fh)
	{
		$self->{'_FH'} = $fh;
	}
	else
	{
		$self->{'_FILENAME'} = $args{'file'} // $args{'filename'}
			or croak "Need filename";
	}

	return $self;
}

sub _prepare
{
	my $self = shift;
	
	return $self if $self->{'document'};
	
	my $namespaces = {
		office  => OFFICE_NS,
		style   => STYLE_NS,
		text    => TEXT_NS,
		table   => TABLE_NS,
		meta    => META_NS,
		number  => NUMBER_NS,
		calcext => CALCEXT_NS,
		};
	
	$self->{'document'} = XML::LibXML->createDocument;
	$self->{'document'}->setDocumentElement(
		$self->{'document'}->createElement('root')
		);
	while (my ($prefix, $nsuri) = each %$namespaces)
	{
		$self->{'document'}->documentElement->setNamespace($nsuri, $prefix, $prefix eq 'office' ? 1 : 0);
	}
	$self->{'document'}->documentElement->setNodeName('office:document-content');
	$self->{'document'}->documentElement->setAttributeNS(OFFICE_NS, 'version', '1.0');
	###### styles
	my $automatic_styles = $self->{'document'}->documentElement
		->addNewChild(OFFICE_NS, 'automatic-styles');
	my $date_style = $automatic_styles->addNewChild(NUMBER_NS, 'date-style');
	$date_style->setAttributeNS(STYLE_NS,'name','N84');
	$date_style->addNewChild(NUMBER_NS,'year')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style->addNewChild(NUMBER_NS,'text')->appendText('-');
	$date_style->addNewChild(NUMBER_NS,'month')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style->addNewChild(NUMBER_NS,'text')->appendText('-');
	$date_style->addNewChild(NUMBER_NS,'day')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style = $automatic_styles->addNewChild(NUMBER_NS, 'date-style');
	$date_style->setAttributeNS(STYLE_NS,'name','N52');
	$date_style->addNewChild(NUMBER_NS,'year')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style->addNewChild(NUMBER_NS,'text')->appendText('-');
	$date_style->addNewChild(NUMBER_NS,'month')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style->addNewChild(NUMBER_NS,'text')->appendText('-');
	$date_style->addNewChild(NUMBER_NS,'day')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style->addNewChild(NUMBER_NS,'text');
	$date_style->addNewChild(NUMBER_NS,'hours')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style->addNewChild(NUMBER_NS,'text')->appendText(':');
	$date_style->addNewChild(NUMBER_NS,'minutes')->setAttributeNS(NUMBER_NS,'style','long');
	$date_style->addNewChild(NUMBER_NS,'text')->appendText(':');
	$date_style->addNewChild(NUMBER_NS,'seconds')->setAttributeNS(NUMBER_NS,'style','long');
	my $time_style = $automatic_styles->addNewChild(NUMBER_NS, 'time-style');
	$time_style->setAttributeNS(STYLE_NS,'name','N41');
	$time_style->addNewChild(NUMBER_NS,'hours')->setAttributeNS(NUMBER_NS,'style','long');
	$time_style->addNewChild(NUMBER_NS,'text')->appendText(':');
	$time_style->addNewChild(NUMBER_NS,'minutes')->setAttributeNS(NUMBER_NS,'style','long');
	$time_style->addNewChild(NUMBER_NS,'text')->appendText(':');
	$time_style->addNewChild(NUMBER_NS,'seconds')->setAttributeNS(NUMBER_NS,'style','long');
	my $ce1 = $automatic_styles->addNewChild(STYLE_NS,'style');
	$ce1->setAttributeNS(STYLE_NS,'name','ce1');
	$ce1->setAttributeNS(STYLE_NS,'family','table-cell');
	$ce1->setAttributeNS(STYLE_NS,'parent-style-name','Default');
	$ce1->setAttributeNS(STYLE_NS,'data-style-name','N84');
	my $ce5 = $automatic_styles->addNewChild(STYLE_NS,'style');
	$ce5->setAttributeNS(STYLE_NS,'name','ce5');
	$ce5->setAttributeNS(STYLE_NS,'family','table-cell');
	$ce5->setAttributeNS(STYLE_NS,'parent-style-name','Default');
	$ce5->setAttributeNS(STYLE_NS,'data-style-name','N41');
	my $ce7 = $automatic_styles->addNewChild(STYLE_NS,'style');
	$ce7->setAttributeNS(STYLE_NS,'name','ce7');
	$ce7->setAttributeNS(STYLE_NS,'family','table-cell');
	$ce7->setAttributeNS(STYLE_NS,'parent-style-name','Default');
	$ce7->setAttributeNS(STYLE_NS,'data-style-name','N52');
	$self->{'body'} = $self->{'document'}->documentElement
		->addNewChild(OFFICE_NS, 'body')
		->addNewChild(OFFICE_NS, 'spreadsheet');
	$self->addsheet($self->{'options'}->{'sheet'} // 'Sheet 1');
	
	return $self;
}

sub addsheet
{
	my ($self, $caption) = @_;

	$self->_open() or return;

	$self->{'tbody'} = $self->{'body'}->addNewChild(TABLE_NS, 'table');

	if (defined $caption)
	{
		$self->{'tbody'}->setAttributeNS(TABLE_NS, 'name', $caption);
	}
	
	return $self;
}

sub _add_prepared_row
{
	my $self = shift;

	my $tr = $self->{'tbody'}->addNewChild(TABLE_NS, 'table-row');
	
	foreach my $cell (@_)
	{
		my $tcell = $tr->addNewChild(TABLE_NS, 'table-cell');
		if ($cell->{content} =~ /^\d\d\d\d-\d\d-\d\d$/) {
			$tcell->setAttributeNS(OFFICE_NS, 'date-value',$cell->{content});
			$tcell->setAttributeNS(OFFICE_NS, 'value-type','date');
			$tcell->setAttributeNS(CALCEXT_NS, 'value-type','date');
			$tcell->setAttributeNS(TABLE_NS,'style-name','ce1');
		}
		elsif ($cell->{content} =~ /^(\d\d):(\d\d):(\d\d)$/) {
			$tcell->setAttributeNS(OFFICE_NS, 'time-value',"PT$1H$2M$3S");
			$tcell->setAttributeNS(OFFICE_NS, 'value-type','time');
			$tcell->setAttributeNS(CALCEXT_NS, 'value-type','time');
			$tcell->setAttributeNS(TABLE_NS,'style-name','ce5');
		}
		elsif ($cell->{content} =~ /^(\d\d\d\d-\d\d-\d\d) (\d\d:\d\d:\d\d)$/) {
			$tcell->setAttributeNS(OFFICE_NS, 'date-value',"$1T$2");
			$tcell->setAttributeNS(OFFICE_NS, 'value-type','date');
			$tcell->setAttributeNS(CALCEXT_NS, 'value-type','date');
			$tcell->setAttributeNS(TABLE_NS,'style-name','ce7');
		}
		elsif ($cell->{content} =~ /^(?:(?i)(?:[-+]?)(?:(?=[.]?[0123456789])(?:[0123456789]{1,3}(?:(?:,?)[0123456789]{3})*)(?:(?:[.])(?:[0123456789]{0,}))?))$/ ) {
		#float English format
		my $value =  $cell->{content} =~ s/,//gr;
			$tcell->setAttributeNS(OFFICE_NS, 'value',$value);
			$tcell->setAttributeNS(OFFICE_NS, 'value-type','float');
			$tcell->setAttributeNS(CALCEXT_NS, 'value-type','float');
			$tcell->setAttributeNS(TABLE_NS,'style-name','Default');
		}
		elsif ($cell->{content} =~ /^(?:(?i)(?:[-+]?)(?:(?=[,]?[0123456789])(?:[0123456789]{1,3}(?:(?:\.?)[0123456789]{3})*)(?:(?:[,])(?:[0123456789]{0,}))?))$/ ) {
		#float spanish format
		my $value =  $cell->{content} =~ s/[.]//gr;
		$value =~ s/,/./;
			$tcell->setAttributeNS(OFFICE_NS, 'value',$value);
			$tcell->setAttributeNS(OFFICE_NS, 'value-type','float');
			$tcell->setAttributeNS(CALCEXT_NS, 'value-type','float');
			$tcell->setAttributeNS(TABLE_NS,'style-name','Default');
		}
		else {
			$tcell->setAttributeNS(OFFICE_NS,'value-type','string')
		}
		
		my $td = $tcell->addNewChild(TEXT_NS, 'p');
		
		my $content = $cell->{'content'};
		$content = sprintf($cell->{'sprintf'}, $content)
			if defined $cell->{'sprintf'};
		
		$td->appendText($content);
		
		if ($cell->{'font_weight'} eq 'bold'
		&&  $cell->{'font_style'} eq 'italic')
		{
			$td->setAttributeNS(TEXT_NS, 'style-name', 'BoldItalic');
		}
		elsif ($cell->{'font_weight'} eq 'bold')
		{
			$td->setAttributeNS(TEXT_NS, 'style-name', 'Bold');
		}
		elsif ($cell->{'font_style'} eq 'italic')
		{
			$td->setAttributeNS(TEXT_NS, 'style-name', 'Italic');
		}
	}
}

sub close
{
	my $self=shift;
	return if $self->{'_CLOSED'};
	$self->{'_FH'}->print( $self->_make_output );
	$self->{'_FH'}->close;
	$self->{'_CLOSED'}=1;
	return $self;
}

sub _make_output
{
	my $self = shift;
	return $self->{'document'}->toString;
}

1;
