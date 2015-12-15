/////////////////////////////////////////////////////////////////
// Proof Details Generator V1.0 CS5
//---------------------------------------------------------------
/////////////////////////////////////////////////////////////////

//Global Variables
doc = activeDocument;

function toTitleCase(str){
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}

function launchHelp(){
  var helpTopic = prompt("Help info: | About | Embroidery | Screen-Print | Update | Exit", "");
  var helpTopicKey = helpTopic.toLowerCase();

  switch(helpTopicKey){
    case "about":
      alert("Proof Details Generator was written by Josh Makar, 2015.\rTo report any issues, bugs, or to suggest improvements, send an email to joshmakar@gmail.com");

      launchHelp();

      break;

    case "embroidery":
      alert("The Embroidery option will output the proof details for embroidery including positions and embelishment sizes.");

      launchHelp();

      break;

    case "screen-print":
      alert("The Screen-Print option will output the proof details for screen-printing including positions, imprint sizes, and ink colors used.\rThe ink colors used will automatically be generated based off of the swatches in your swatch pallet.");

      launchHelp();

      break;

    case "update":
      alert("To ensure Proof Details Generator is up-to-date, visit:\rhttp://github.com/joshmakar/proof-details-generator");

      launchHelp();

      break;

    case "exit":
      break;

    default:
      alert("Please type a valid response or type 'exit'.");
      launchHelp();
  }
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
    textRef.contents = "Colors \rUsed";
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
    var productionType = prompt("What is the production type? Or type /help","Screen-Print, Embroidery, Digital, Heatpress, Vinyl, or Other");
    var productionTypeKey = productionType.toLowerCase();
    var productionTypeKey = productionTypeKey.substring(0, 3);

    var p = 0;
    var n = true;
    var position = [];

    var y = -640;
    var x = 30;
    var w = 100;
    var h = 10;
    var margin = 5;

    var white = new GrayColor()
    var noColor = new NoColor();
    var white.gray = 0;

    var pointTextRef = function(x, y, contents){
        TextRef = doc.textFrames.add();
        TextRef.textRange.characterAttributes.size = 10;
        TextRef.position = [x, y];
        TextRef.contents = contents;
    }

	var productionTypeBG = function(){
		TextRef = doc.pathItems.rectangle(y,x - 2, contents.length * 5.5,h);
	    TextRef.fillColor = white;
	    TextRef.strokeColor = noColor;
	    TextRef.strokeWidth = "0";
	}

	if(productionTypeKey != "/he"){
		newGroup = doc.layers.add(); //var newGroup = doc.groupItems.add();
		newGroup.name = "Proof Details";
	}

    switch (productionTypeKey) {
        case "scr":
        	contents = "Screen-Print";
        	productionTypeBG();

            pointTextRef(x, y, contents);

            while(p < 4 && n == true){
                position[p] = prompt("What is the Print Position","Full Front, Full Back, Left Chest, Etc.");

                pointTextRef(x + (w * p), y - (h + margin), toTitleCase(position[p]) + '\rTBD"w x TBD"h');

                p++;

                if(p < 4){
                    n = confirm('Add another position?');
                }
            }

            createSwatches();
            
            break;
      
        case "dig":
            var productName = prompt("What is the Product Name?","Banner, Yard Sign, A-Frame, Etc.");
            var productWidth = prompt("What is the Product Width",'6\', 24", Etc.');
            var productHeight = prompt("What is the Product Height",'3\', 18", Etc.');

            contents = "Digital Printing";
        	productionTypeBG();

            pointTextRef(x, y, contents);

            pointTextRef(x, y - (h + margin), toTitleCase(productName) + '\r' + productWidth + 'w x ' + productHeight + 'h');

            break;

        case "emb":
            var productName = prompt("What is the Product Type?","Polos, Hats, Etc.");

            contents = "Embroidery";
        	productionTypeBG();

            pointTextRef(x, y, contents);

            while(p < 4 && n == true){
              position[p] = prompt("What is the Embroidery Position","Left Chest, Right Chest, Center Front, Etc.");

              pointTextRef(x + (w * p), y - (h + margin), toTitleCase(position[p]) + '\rTBD"w x TBD"h');

              p++;
              
              if(p < 4){
                  n = confirm('Add another position?');
              }

            }

            break;

        case "hea":
            var productName = prompt("What is the Product Type?","Shirts, Cinch Bags, Etc.");

            contents = "Heatpress";
        	productionTypeBG();

            pointTextRef(x, y, contents);

            while(p < 4 && n == true){
              position[p] = prompt("What is the Heatpress Position","Left Chest, Full Back, Right Sleeve, Etc.");

              pointTextRef(x + (w * p), y - (h + margin), toTitleCase(position[p]) + '\rTBD"w x TBD"h');

              p++;
              
              if(p < 4){
                  n = confirm('Add another position?');
              }

            }

            break;

        case "vin":
            var productWidth = prompt("What is the Product Width",'6", 24", Etc.');
            var productHeight = prompt("What is the Product Height",'3", 18", Etc.');

            contents = "Cut Vinyl";
        	productionTypeBG();

            pointTextRef(x, y, contents);

            pointTextRef(x, y - (h + margin), productWidth + 'w x ' + productHeight + 'h');

            break;

        case "/he":
          launchHelp();

          break;

        default:
            contents = toTitleCase(productionType);
        	productionTypeBG();

            pointTextRef(x, y, contents);

            pointTextRef(x, y - (h + margin), "Product" + '\rTBD"w x TBD"h');
    }
}

outputProductionDetails();