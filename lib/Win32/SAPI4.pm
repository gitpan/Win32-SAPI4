#!/usr/bin/perl
package Win32::SAPI4;
use vars qw ($VERSION);
$VERSION     = 0.01;

package Win32::SAPI4::VoiceText;
use Win32::OLE qw( EVENTS );

sub new
{
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    my $self = Win32::OLE->new('{2398E32F-5C6E-11D1-8C65-0060081841DE}') || die "Can't start VoiceText";
    return $self;
}

package Win32::SAPI4::VoiceCommand;
use Win32::OLE qw( EVENTS );

sub new
{
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    my $self = Win32::OLE->new('{66523042-35FE-11D1-8C4D-0060081841DE}') || die "Can't start VoiceText";
    return $self;
}

package Win32::SAPI4::VoiceDictation;
use Win32::OLE qw( EVENTS );

sub new
{
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    my $self = Win32::OLE->new('{582C2191-4016-11D1-8C55-0060081841DE}') || die "Can't start VoiceText";
    return $self;
}

package Win32::SAPI4::VoiceTelephony;
use Win32::OLE qw( EVENTS );

sub new
{
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    my $self = Win32::OLE->new('{FC9E740F-6058-11D1-8C66-0060081841DE}') || die "Can't start VoiceText";
    return $self;
}

package Win32::SAPI4::DirectSpeechRecognition;
use Win32::OLE qw( EVENTS );

sub new
{
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    my $self = Win32::OLE->new('{4E3D9D1F-0C63-11D1-8BFB-0060081841DE}') || die "Can't start VoiceText";
    return $self;
}

package Win32::SAPI4::DirectSpeechSynthesis;
use Win32::OLE qw( EVENTS );

sub new
{
    Win32::OLE->Initialize(Win32::OLE::COINIT_MULTITHREADED);
    my $self = Win32::OLE->new('{EEE78591-FE22-11D0-8BEF-0060081841DE}') || die "Can't start VoiceText";
    return $self;
}

=pod

=head1 NAME

Win32::SAPI4

=head1 SYNOPSIS

    use Win32::SAPI4;
    
    my $vt   = Win32::SAPI4::VoiceText->new();
    my $dss  = Win32::SAPI4::DirectSpeechSynthesis->new();
    my $dsr  = Win32::SAPI4::DirectSpeechRecognition->new();
    my $vtel = Win32::SAPI4::VoiceTelephony->new();    
    my $vd   = Win32::SAPI4::VoiceDictation->new();
    my $vc   = Win32::SAPI4::VoiceCommand->new();

=head1 DESCRIPTION

This module is a simple interface to the Microsoft Speech API 4.0.
It just offers 6 constructors to the different parts of this API.
This documentation won't offer the complete documentation for it, just
download the Microsoft Speech API 4.0 SDK and read the Visual Basic
documentation. It's all there.

=head1 PREREQUISITES

The Microsoft Speech API 4.0. It can be downloaded for free from 
http://www.microsoft.com/speech (go to the 'old versions' and 
find the 4.0 version)

=head1 USAGE

See the Microsoft Speech API 4.0 Visual Basic documentation for 
all methods, properties and events available.

=head1 SUPPORT

The Microsoft SAPI 4.0 SDK is supported on news://microsoft.public.speech_tech.sdk
You can email the author for support on this module.

=head1 AUTHOR

	Jouke Visser
	jouke@cpan.org
	http://jouke.pvoice.org

=head1 COPYRIGHT

Copyright (c) 2003 Jouke Visser. All rights reserved.
This program is free software; you can redistribute
it and/or modify it under the same terms as Perl itself.

The full text of the license can be found in the
LICENSE file included with this module.

=head1 SEE ALSO

perl(1), Microsoft Speech API 4 documentation.

=cut

1;