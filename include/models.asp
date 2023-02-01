this.Header=function Header()
{
    this.SetFont('Arial','B',15);
    this.Cell(80);
    this.Cell(30,10,'Title',1,0,'C');
    this.Ln(20);
} 


this.AcceptPageBreak = function AcceptPageBreak()
	{
		if(this.y0+this.rowheight>this.PageBreakTrigger)
			return true;
		var x;
		x=this.leftmargin;
		if(this.maxheight<this.PageBreakTrigger-this.y0)
			this.maxheight=this.PageBreakTrigger-this.y0;
##foreach Fields as @f filter @f.bExportPage order @f.nExportPageOrder##
		this.Rect(x,this.y0,this.colwidth["##@f.strName s##"],this.maxheight);
		x+=this.colwidth["##@f.strName s##"];
##endfor##
		this.maxheight = this.rowheight;
//	draw frame	
		return true;
	}

this.Header=function Header()
{
	this.SetFillColor(192);
	this.SetX(this.leftmargin);
##foreach Fields as @f filter @f.bExportPage order @f.nExportPageOrder##
	this.Cell(this.colwidth["##@f.strName s##"],this.rowheight,"##@f.strLabel s##",1,0,'C',1);
##endfor##
	this.Ln(this.rowheight);
	this.y0=this.GetY();
}

this.GetColWidth = function GetColWidth(name)
{
	return this.colwidth[name];
}

this.leftmargin=5;
pagewidth=200;
pageheight=290;
this.rowheight=5;


defwidth=pagewidth/##Fields[bExportPage].len##;
this.colwidth=new Array();
##foreach Fields as @f filter @f.bExportPage order @f.nExportPageOrder##
    this.colwidth["##@f.strName s##"]=defwidth;
##endfor##
