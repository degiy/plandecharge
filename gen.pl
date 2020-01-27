#!/usr/bin/perl

use List::Util qw(sum);

$verbose=1;

# generation des onglets du plan de charge en fonction de la feuille config export�e en csv par le vba

@tres=();
@tact=();
@files=();
$reload=0;
$nba=1;

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
	$res=~s/ *$//;
	$res=~s/^ *//;
	push @tres,$res;
	$htmres{$res}=[@mres];
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
	$act=~s/ *$//;
	$act=~s/^ *//;
	push @tact,$act;
	$htact{$act}{'BC'}=shift @t;
	$htact{$act}{'code'}=shift @t;
	$htact{$act}{'mdeb'}=shift @t;
	$htact{$act}{'mfin'}=shift @t;
	$htact{$act}{'hjtot'}=shift @t;
	$htact{$act}{'hjdjc'}=shift @t;
	$htact{$act}{'id'}=$nba++;
	@qui=sort @t;
	foreach $r (@qui)
	{
	    unless ($r eq '')
	    {
		push @{$htares{$r}},$act;
	    }
	}
    }
    # fichier supplementaire a charger sous forme d'onglet suppl.
    elsif (/^F/)
    {
	s/^F,//;
	push @files,$_;
    }
}
close F;

# post traitements
@tres=sort @tres;
@tact=sort @tact;
$nbmois=$#mois+1;
@zm=(0)x$nbmois;

@xlet=split /,/,"A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z,AA,AB,AC,AD,AE,AF,AG,AH,AI,AJ,AK,AL,AM,AN,AO,AP,AQ,AR,AS,AT,AU,AV,AW,AX,AY,AZ";

# faut-il recharger les charges declarees pour les ressource d'un tableur precedent ?
goto gen unless  ($#ARGV==0) && ( $ARGV[0] eq '-r' );

open(F,"saved_ressources.csv") or die;
$reload=1;
$etat=0;
while (<F>)
{
    s/\r//gs;
    s/\n//gs;
    # liste des mois
    if (/^mois,.*,total/)
    {
	# recup de la ligne des mois (par forcement les memes qu'on a dans le nouveau plan de charge)
	@omois=split /,/,$_;
	shift @omois; pop @omois;
	print "old mois : ",join(',',@omois),"\n" if $verbose;
    }
    elsif (/^,*$/)
    {
	# ligne vide (donc nouvelle ressource, pas forcement presente dans le nouveau plan de charge)
	$etat=1;
    }
    elsif (/^total/)
    {
	# fin du bloc de conso
	$etat=0;
    }
    elsif ($etat==1)
    {
	# ligne dispo ressource par mois (car on est dans l'etat 1)
	# pierre,5,5,5,5,5,5,5,5,5,5,5,5,60
	print "1 : $_ \n" if $verbose;
	@rmois=split /,/,$_;
	$r=shift @rmois;
	pop @rmois;
	print "  -> ",join(',',@rmois),"\n" if $verbose;
	$etat=2;
    }
    elsif (($etat==2)&&(exists $htares{$r}))
    {
	# ligne d'affectation de jours d'une ressource a une activite
	# archi,1,1,1,1,1,1,1,1,1,1,1,1,12
	print "2 : $_ \n" if $verbose;
	@ra=split /,/,$_;
	$a=shift @ra;
	if (exists $htact{$a})
	{
	    pop @ra;
	    print "    -> ($a) ",join(',',@ra),"\n" if $verbose;
	    for($i=0;$i<$#ra;$i++)
	    {
		$prev=$ra[$i];
		$m=$omois[$i];
		$htoldra{$r}{$a}{$m}=$prev;
		print "     $r/$a/$m=$prev\n" if $verbose>1;
	    }
	}
    }
}

close F;

 gen: ;

# generation des styles
open (FS,">styles.csv") or die;

# generation du fichier csv des ressources (celui ou on va choisir combien la ressource R impute chaque mois sur ces differentes activites)
open (F,">ressources_raw.csv") or die;
$y=1;
print F "mois,",join(',',@mois),",total\n\n";
$y+=2;
foreach $r (@tres)
{
    @max=@{$htmres{$r}};
    #$tot=sum @max;
    print F "$r,",join(',',@max),",=SOMME(B$y:",$xlet[$nbmois],"$y)\n";
    print FS "color,ressources,A$y:A$y,&H99ccff\n";
    $y++;
    $y0=$y;
    foreach $a (@{$htares{$r}})
    {
	print F "$a";
	if ($reload)
	{
	    for($x=0;$x<$nbmois;$x++)
	    {
		$charge=0;
		$charge=$htoldra{$r}{$a}{$mois[$x]} if exists $htoldra{$r}{$a}{$mois[$x]};
		print F ",$charge";
	    }
	}
	else
	{
	    print F ',',join(',',@zm);
	}
	print F ",=SOMME(B$y:",$xlet[$nbmois],$y,"),=activites!%%",$htact{$a}{'id'},"%%\n";
	$hty{$a}{$r}=$y;
	$y++;	
    }
    print FS "color,ressources,B$y0:",$xlet[$nbmois],$y-1,",&Hff9966\n";
    print F "total(j)";
    for($x=1;$x<=$nbmois+1;$x++)
    {
	$let=$xlet[$x];
	print F ",=SOMME($let$y0:$let",$y-1,")";
    }
    print F "\n";
    $y++;
    print F "total(%)";
    for($x=1;$x<=$nbmois+1;$x++)
    {
	$let=$xlet[$x];
	print F ",=$let",$y-1,"/$let",$y0-1;
    }
    print FS "pct,ressources,B$y:",$xlet[$nbmois+1],"$y\n";
    print F "\n\n";
    $y+=2;
}
close F;
# rechargement pour remplacements
open (F,"ressources_raw.csv") or die;
read F,$rsfile,1000000;
close F;

# generation du fichier csv des activites (recap des estimations de conso par rapport au volume de l'activite)
open (F,">activites.csv") or die;
$y=1;
foreach $a (@tact)
{
    # les ressource sur l'activite
    @res=sort keys %{$hty{$a}};
    $nbres=$#res+1;
    # info generales sur l'activite
    print F "activite,BC,code,deb,fin,total(j),deja(j),rafi(j),raf(j),raf(%)\n";
    $y++;
    #$raf=$htact{$a}{'hjtot'}-$htact{$a}{'hjdjc'};
    $raf="=F$y-G$y";
    print F "$a,",$htact{$a}{'BC'},',',$htact{$a}{'code'},',',$htact{$a}{'mdeb'},',',$htact{$a}{'mfin'},',',$htact{$a}{'hjtot'},',',$htact{$a}{'hjdjc'},',',$raf,',';
    # raf j et % (par rapport au tableau de la conso ressource juste en dessous)
    print F "=H$y-",$xlet[$nbmois+1],$y+$nbres+2,",=I$y/F$y\n";
    print FS "pct,activites,J$yJ$y\n";
    print FS "color,activites,A$y:A$y,&Hcc99FF\n";
    print FS "color,activites,I$y:J$y,&HFFcc99\n";
    $y++;
    print F "ressource,",join(',',@mois),",total(j),total(%)\n";
    $y++;
    $y1=$y;
    # pour chaque ressource, recopie des imputs de la feuille des ressources
    foreach $r (@res)
    {
	print F "$r";
	for($x=1;$x<=$nbmois+1;$x++)
	{
	    $let=$xlet[$x];
	    $y0=$hty{$a}{$r};
	    print F ",=ressources!$let$y0";
	}
	# pct du raf initial
	print F ",=$let$y/H",$y1-2,"\n";
	$y++;	
    }
    # total
    print F "total(j)";
    for($x=1;$x<=$nbmois+2;$x++)
    {
	$let=$xlet[$x];
	$somme="=SOMME($let$y1:$let".($y-1).")";
	$somme="0" if $nbres==0;
	print F ",$somme";
    }
    print FS "pct,activites,",$xlet[$nbmois+2],$y1,':',$xlet[$nbmois+2],$y,"\n";  
    print FS "color,activites,B$y:",$xlet[$nbmois+1],"$y,&Hcccccc\n";  
    print F "\n\n";
    $pats='%%'.$htact{$a}{'id'}.'%%';
    $patd='I'.($y1-2);
    $rsfile=~s/$pats/$patd/sg;
    $y+=2;
}
close F;
close FS;

# on ecrit le fichier de ressources avec les placeholders de % d'activite globale remplaces 
open (F,">ressources.csv") or die;
print F $rsfile;
close F;
