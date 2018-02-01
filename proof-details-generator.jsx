/////////////////////////////////////////////////////////////////
//  Proof Details Generator CS5
    var PDGVersion = "V1.2.1";
//---------------------------------------------------------------
/////////////////////////////////////////////////////////////////

// Global Variables
doc = activeDocument;
myfont = "DroidSans";
myfontMono = "DroidSansMono";

function toTitleCase(str){
    return str.replace(/\w\S*/g, function(txt){return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();});
}

function launchHelp(){
  var helpTopic = "";
  var helpTopic = prompt("Help info:\rAbout | Production Type | Tips | Update | Exit", "");
  if(helpTopic == null){
    outputProductionDetails();
  } else {

    var helpTopicKey = helpTopic.toLowerCase();

    switch(helpTopicKey){
      case "about":
        alert("Proof Details Generator " + PDGVersion + ". written by Josh Makar, 2015.\rTo report any issues, bugs, or to suggest improvements, send an email to joshmakar@gmail.com");

        launchHelp();

        break;

      case "production type":
        alert("The Production Type is the type of printing, embroidery, or other embelishment. Each different production type will output a specific set of details including imprint sizes, ink colors, and product names.");

        launchHelp();

        break;

      case "tips":
        alert("You only need to type the first three characters of the default Production Type, e.g. scr = Screen-Print.\r\rIf you'd like to output the swatch colors only, type 'colors'.\r\rTo disable title case formating, prepend your entry with a \\, e.g. \\PVC Sign.");

        launchHelp();

        break;

      case "update":
        alert("You are running Proof Details Generator " + PDGVersion + "\rTo ensure Proof Details Generator is up-to-date, visit:\rhttp://github.com/joshmakar/proof-details-generator");

        launchHelp();

        break;

      case "exit":
        break;

      default:
        alert("Please type a valid response or type 'exit'.");
        launchHelp();
    }
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
    try{
      textRef.textRange.characterAttributes.textFont=app.textFonts.getByName(myfont);
    }
    catch(e){alert("Please install the Google font: " + myfont)}
    textRef.contents = "Colors\rUsed";
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
            rectRef.strokeDashes = [];
            textRectRef =  doc.pathItems.rectangle(y- t_v_pad,x+ t_h_pad, w-(2*t_h_pad),h-(2*t_v_pad));
            textRef = doc.textFrames.areaText(textRectRef);
            textRef.textRange.characterAttributes.size = 7;
            try{
              textRef.textRange.characterAttributes.textFont=app.textFonts.getByName(myfont);
            }
            catch(e){alert("Please install the Google font: " + myfont)}
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
    var productionType = "";
    var productionType = prompt("What is the production type? Or type /help","Screen-Print, Embroidery, Digital, Heatpress, Vinyl, or Other");
    if(productionType != null){
      var productionTypeKey = productionType.toLowerCase();
      var productionTypeKey = productionTypeKey.substring(0, 3);
    }

    var p = 0;
    var n = true;
    var position = [];
    var designID = [];

    var y = -640;
    var x = 30;
    var w = 100;
    var h = 10;
    var margin = 5;

    white = new GrayColor();
    noColor = new NoColor();
    white.gray = 0;

    var pointTextRef = function(x, y, contents){
        TextRef = doc.textFrames.add();
        try{  
          TextRef.textRange.characterAttributes.textFont=app.textFonts.getByName(myfont);  
        }  
        catch(e){alert("Please install the Google font: " + myfont)} 
        TextRef.textRange.characterAttributes.size = 10;
        TextRef.position = [x, y];
        TextRef.contents = contents;
    }

    var pointTextRefMono = function(x, y, contents){
        TextRef = doc.textFrames.add();
        try{
          TextRef.textRange.characterAttributes.textFont=app.textFonts.getByName(myfontMono);
        }
        catch(e){alert("Please install the Google font: " + myfontMono)}
        TextRef.textRange.characterAttributes.size = 10;
        TextRef.position = [x, y];
        TextRef.contents = contents;
    }

	var productionTypeBG = function(){
        whiteBox = doc.pathItems.rectangle(y,x - 2, contents.length * 5.876 + 6,h);
        whiteBox.fillColor = white;
        whiteBox.strokeColor = noColor;
        whiteBox.strokeWidth = "0";
	}

	if(productionTypeKey != "/he" && productionTypeKey != "/re" && productionTypeKey != "col" && productionTypeKey != null){
    try {  
        doc.layers.getByName( 'Proof Details' ).remove();  
    } catch (e) {};

		newGroup = doc.layers.add(); //var newGroup = doc.groupItems.add();
		newGroup.name = "Proof Details";
	}

  if(productionType != null){
    switch (productionTypeKey) {
        case "scr":
          	contents = "Screen-Print";
          	productionTypeBG();

            pointTextRefMono(x, y, contents);

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

            if (productName.charAt(0) == "\\") {
            	productName = productName.substr(1);
        	} else {
            	productName = toTitleCase(productName);
            }

            contents = "Digital Printing";
            productionTypeBG();

            pointTextRefMono(x, y, contents);

            pointTextRef(x, y - (h + margin), productName + '\r' + productWidth + 'w x ' + productHeight + 'h');

            break;

        case "emb":
            contents = "Embroidery";
            productionTypeBG();

            pointTextRefMono(x, y, contents);

            while(p < 4 && n == true){
              position[p] = prompt("What is the Embroidery Position","Left Chest, Right Chest, Center Front, Etc.");
              designID[p] = prompt("What is the Design ID for this Position?","");

              pointTextRef(x + (w * p), y - (h + margin), toTitleCase(position[p]) + '\rTBD"w x TBD"h\r' + (designID[p] ? 'Design ID: ' + designID[p] : ''));

              p++;
              
              if(p < 4){
                  n = confirm('Add another position?');
              }

            }

            break;

        case "hea":
            contents = "Heatpress";
        	productionTypeBG();

            pointTextRefMono(x, y, contents);

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

            pointTextRefMono(x, y, contents);

            pointTextRef(x, y - (h + margin), productWidth + 'w x ' + productHeight + 'h');

            break;

        case "col":

          newGroup = doc.layers.add();
          newGroup.name = "Colors Used";

          createSwatches();

          break;

        case "/he":
          launchHelp();

          break;

        case "/re":
          try {  
              doc.layers.getByName( 'Proof Details' ).remove();  
          } catch (e) {};

          break;

        default:
            contents = toTitleCase(productionType);
        	productionTypeBG();

            pointTextRefMono(x, y, contents);

            pointTextRef(x, y - (h + margin), "Product" + '\rTBD"w x TBD"h');
    }
  }
}

outputProductionDetails();