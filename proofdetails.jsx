﻿/////////////////////////////////////////////////////////////////
// Render Proof Details V1.0 CS5
//---------------------------------------------------------------
/////////////////////////////////////////////////////////////////

doc = activeDocument;

var newGroup = doc.layers.add(); //var newGroup = doc.groupItems.add();
newGroup.name = "Proof Details";
newGroup.move( doc, ElementPlacement.PLACEATBEGINNING );

function toTitleCase(str){
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}

function createSwatches(){

    swatches = doc.swatches,
    //cols = 9,
    displayAs = "CMYKColor",  //or "RGBColor"
    rectRef=null,
    textRectRef=null,
    textRef=null,
    rgbColor=null,
    w=66;
    h=26,
    h_pad = 5,
    v_pad = 5,
    t_h_pad = 5,
    t_v_pad = 5,
    x=null,
    y=null,
    black = new GrayColor(),
    white = new GrayColor()
    ;

    black.gray = 100;
    white.gray = 0;

    textRef = doc.textFrames.areaText(doc.pathItems.rectangle(-700, 30, 50, 30));
    textRef.textRange.characterAttributes.size = 10;
    textRef.contents = "Ink Colors Used";
    textRef.move( newGroup, ElementPlacement.PLACEATBEGINNING );

    for(var c=2,len=swatches.length;c<len && c<9;c++)
    {
            var swatchGroup = doc.groupItems.add();
            swatchGroup.name = swatches[c].name;
           
            x= (w+h_pad)*(c-1)+20;
            y=-700;
            rectRef = doc.pathItems.rectangle(y,x, w,h);
            rgbColor = swatches[c].color;
            rectRef.fillColor = rgbColor;
            rectRef.strokeColor = black;
            rectRef.strokeWidth = "1";
            textRectRef =  doc.pathItems.rectangle(y- t_v_pad,x+ t_h_pad, w-(2*t_h_pad),h-(2*t_v_pad));
            textRef = doc.textFrames.areaText(textRectRef);
            textRef.textRange.characterAttributes.size = 8;
            textRef.contents = swatches[c].name;
            textRef.textRange.fillColor = is_dark(swatches[c].color)? white : black;
            //
            rectRef.move( swatchGroup, ElementPlacement.PLACEATBEGINNING );     
            textRef.move( swatchGroup, ElementPlacement.PLACEATBEGINNING );
            swatchGroup.move( newGroup, ElementPlacement.PLACEATEND );
    }

    function getColorValues(color)
    {
            if(color.typename)
            {
                switch(color.typename)
                {
                    case "CMYKColor":
                        if(displayAs == "CMYKColor"){
                            return ([Math.floor(color.cyan),Math.floor(color.magenta),Math.floor(color.yellow),Math.floor(color.black)]);}
                        else
                        {
                            color.typename="RGBColor";
                            return  [Math.floor(color.red),Math.floor(color.green),Math.floor(color.blue)] ;
                           
                        }
                    case "RGBColor":
                       
                       if(displayAs == "CMYKColor"){
                            return rgb2cmyk(Math.floor(color.red),Math.floor(color.green),Math.floor(color.blue));
                       }else
                        {
                            return  [Math.floor(color.red),Math.floor(color.green),Math.floor(color.blue)] ;
                        }
                    case "GrayColor":
                        if(displayAs == "CMYKColor"){
                            return rgb2cmyk(Math.floor(color.gray),Math.floor(color.gray),Math.floor(color.gray));
                        }else{
                            return [Math.floor(color.gray),Math.floor(color.gray),Math.floor(color.gray)];
                        }
                    case "SpotColor":
                        return getColorValues(color.spot.color);
                }    
            }
        return "Non Standard Color Type";
    }
    function rgb2cmyk (r,g,b) {
     var computedC = 0;
     var computedM = 0;
     var computedY = 0;
     var computedK = 0;

     //remove spaces from input RGB values, convert to int
     var r = parseInt( (''+r).replace(/\s/g,''),10 ); 
     var g = parseInt( (''+g).replace(/\s/g,''),10 ); 
     var b = parseInt( (''+b).replace(/\s/g,''),10 ); 

     if ( r==null || g==null || b==null ||
         isNaN(r) || isNaN(g)|| isNaN(b) )
     {
       alert ('Please enter numeric RGB values!');
       return;
     }
     if (r<0 || g<0 || b<0 || r>255 || g>255 || b>255) {
       alert ('RGB values must be in the range 0 to 255.');
       return;
     }

     // BLACK
     if (r==0 && g==0 && b==0) {
      computedK = 1;
      return [0,0,0,1];
     }

     computedC = 1 - (r/255);
     computedM = 1 - (g/255);
     computedY = 1 - (b/255);

     var minCMY = Math.min(computedC,
                  Math.min(computedM,computedY));
     computedC = (computedC - minCMY) / (1 - minCMY) ;
     computedM = (computedM - minCMY) / (1 - minCMY) ;
     computedY = (computedY - minCMY) / (1 - minCMY) ;
     computedK = minCMY;

     return [Math.floor(computedC*100),Math.floor(computedM*100),Math.floor(computedY*100),Math.floor(computedK*100)];
    }

    function is_dark(color){
           if(color.typename)
            {
                switch(color.typename)
                {
                    case "CMYKColor":
                        return (color.black>50 || (color.cyan>50 &&  color.magenta>50)) ? true : false;
                    case "RGBColor":
                        return (color.red<100  && color.green<100 ) ? true : false;
                    case "GrayColor":
                        return color.gray > 50 ? true : false;
                    case "SpotColor":
                        return is_dark(color.spot.color);
                    
                    return false;
                }
            }
    }
}

function outputProductionDetails(){

    var productionType = prompt("What is the production type?","Screen-Print, Embroidery, Digital, Heatpress, Vinyl, or Other");
    var productionTypeKey = productionType.toLowerCase();
    var productionTypeKey = productionTypeKey.substring(0, 3);

    var p = 0;
    var n = true;
    var position = [];

    y = -640;
    x = 30;
    w = 100;
    h = 10;
    margin = 5;

    switch (productionTypeKey) {
      case "scr":
        textRef = doc.textFrames.add();
        textRef.textRange.characterAttributes.size = 10;
        textRef.position = [x, y];
        textRef.contents = "Screen-Printing";
        textRef.move( newGroup, ElementPlacement.PLACEATBEGINNING );

        while(p < 4 && n == true){
            position[p] = prompt("What is the Print Position","Full Front, Full Back, Left Chest, Etc.");

            rectRef = doc.pathItems.rectangle(y - (h + margin), x + (w * p), w, h * 2)
            textRef = doc.textFrames.areaText(rectRef);
            textRef.textRange.characterAttributes.size = 10;
            textRef.contents = toTitleCase(position[p]) + '\rTBD"w x TBD"h';
            textRef.move( newGroup, ElementPlacement.PLACEATBEGINNING );

            p++;
            n = confirm('Add another position?');
        }

        createSwatches();
        break;
      case "dig":
        var productName = prompt("What is the Product Name?","Banner, Yard Sign, A-Frame, Etc.");
        var productWidth = prompt("What is the Product Width",'6\', 24", Etc.');
        var productHeight = prompt("What is the Product Height",'3\', 18", Etc.');

        textRef = doc.textFrames.add();
        textRef.textRange.characterAttributes.size = 10;
        textRef.position = [x, y];
        textRef.contents = "Digital Printing";

        textRef = doc.textFrames.add();
        textRef.textRange.characterAttributes.size = 10;
        textRef.position = [x, y - (h + margin)];
        textRef.contents = toTitleCase(productName) + '\r' + productWidth + 'w x ' + productHeight + 'h';
        textRef.move( newGroup, ElementPlacement.PLACEATBEGINNING );
        break;
      case "emb":
        var productName = prompt("What is the Product Type?","Polos, Hats, Etc.");

        textRectRefProductionType =  doc.pathItems.rectangle(y, x, w, h);
        textRefProductionType = doc.textFrames.areaText(textRectRefProductionType);
        textRefProductionType.textRange.characterAttributes.size = 10;
        textRefProductionType.contents = "Embroidery";
        textRefProductionType.move( newGroup, ElementPlacement.PLACEATBEGINNING );

        while(p < 4 && n == true){
          position[p] = prompt("What is the Embroidery Position","Left Chest, Right Chest, Center Front, Etc.");

          textRef = doc.textFrames.areaText(doc.pathItems.rectangle(y - (h + margin), x + (w * p), w, h * 2));
          textRef.textRange.characterAttributes.size = 10;
          textRef.contents = toTitleCase(position[p]) + '\rTBD"w x TBD"h';
          textRef.move( newGroup, ElementPlacement.PLACEATBEGINNING );

          p++;
          n = confirm('Add another position?');
        }
        break;
      case "hea":
                textRefProductionType.textRange.characterAttributes.size = 10;
        break;
      case "hid":
        textRef = doc.textFrames.add();
        textRef.textRange.characterAttributes.size = 10;
        textRef.position = [200,200];
        textRef.contents = "Screen-Printing";
        break;
      default:
        textRectRefProductionType =  doc.pathItems.rectangle(y, x, w, h);
        textRefProductionType = doc.textFrames.areaText(textRectRefProductionType);
        textRefProductionType.textRange.characterAttributes.size = 10;
        textRefProductionType.contents = toTitleCase(productionType);
        textRefProductionType.move( newGroup, ElementPlacement.PLACEATBEGINNING );

        textRectRefproductName =  doc.pathItems.rectangle(y - (h + margin), x, w, h * 2);
        textRefproductName = doc.textFrames.areaText(textRectRefproductName);
        textRefproductName.textRange.characterAttributes.size = 10;
        textRefproductName.contents = "Product" + '\rTBD"w x TBD"h';
        textRefproductName.move( newGroup, ElementPlacement.PLACEATBEGINNING );
    }
}

outputProductionDetails();