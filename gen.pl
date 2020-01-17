#!/usr/bin/perl

$verbose=1;

# generation des onglets du plan de charge en fonction de la feuille config exportée en csv par le vba

@tres=();
@tact=();
@files=();

# lecture de la config
open (F,'config.csv') or die "pas de fichier config.csv a lire";
while(<F>)
{
    # dos CR + LF
    s/\r//gs;
    s/\n//gs;
    # liste des mois
    if (/^M/)
    {
	# M,,2001,2002,2003,2004,2005,2006,2007,2008,2009,2010,2011,2012
	@mois=split /,/,$_;
	shift @mois; shift @mois;
	print "mois : ",join(',',@mois),"\n" if $verbose;
    }
    # ressource
    elsif (/^R/)
    {
	# R,paul,10,15,20,20,20,20,0,0,20,20,10,20
	@mres=split /,/,$_;
	shift @mres;
	$res=shift @mres;
	push @tres,$res;
	$htmres{$res}=@mres;
	$htares{$res}=[];
	print "ressource $res : ",join (',',@mres),"\n" if $verbose;
    }
    # activite
    elsif (/^A/)
    {
	# A,esb,BC-2001,dev,2001,2006,50,10,cyril,pierre,paul,,,
	@t=split /,/,$_;
	shift @t;
	$act=shift @t;
	push @tact,$act;
	$htact{$act}{'BC'}=shift @t;
	$htact{$act}{'code'}=shift @t;
	$htact{$act}{'mdeb'}=shift @t;
	$htact{$act}{'mfin'}=shift @t;
	$htact{$act}{'hjtot'}=shift @t;
	$htact{$act}{'hjdjc'}=shift @t;
	@qui=sort @t;
	@qui2=();
	foreach $r (@qui)
	{
	    unless ($r eq '')
	    {
		push @{$htares{$r}},$act;
		push @qui2,$r;
	    }
	}
	$htact{$act}{'qui'}=@qui2;
	print "activite $act : ",join(',',@qui2),"\n" if $verbose;
    }
    # fichier supplementaire a charger sous forme d'onglet suppl.
    elsif (/^F/)
    {
	s/^F,//;
	push @files,$_;
    }
}
close F;

# generation du fichier csv des ressources


