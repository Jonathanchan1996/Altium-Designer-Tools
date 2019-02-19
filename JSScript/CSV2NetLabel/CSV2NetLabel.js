/**
 ******************************************************************************
 * @project    CSV To Altium Designer Sch Net label
 * @author     Jonathan Chan @ HKUST Robotics Team
 * @Edited     Jonathan Chan
 * @Date       2019/02/11
 ******************************************************************************
 * @Functions
 *     Generate Net in altium designer schematic from CSV
 *
 * @attention
 *
 *
 ******************************************************************************
 */
//User editable variables
var filePath = "D:\\Jonathan\\pinPort.csv";	//File Path
var genLocX = 1000; //Net generated location of X
var genLocY = 500; //Net generated location of Y
var schMaxY = 6000; //paper height
var wireLength = 800;
var spacingY = 100; //space between two wires in Y axis

//Function variables. Please don't edit below
var pinName = new Array();
//var csvData = new Array([]);

function main() {
    /* Doc check */
    if (checkDocType()) {
        return; //This doc is not schematic. End this programme
    }

    /* run main function */
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var f = fso.OpenTextFile(filePath, 1);
    var s = "";

    for (var i = 0; !f.AtEndOfStream; i++) {
        s = f.ReadLine();
        pinName[i] = s;
        ss = s.split("\n");
    }
    f.Close();
    for (var i = 0; i < pinName.length; i++) { //create Net objects
        createNetLabel(genLocX, schMaxY - (genLocY + spacingY * i), eRotate0, pinName[i]);
        createWire(genLocX, schMaxY - (genLocY + spacingY * i), genLocX + wireLength, schMaxY - (genLocY + spacingY * i));
    }
}

function checkDocType(dummy) { //"Run script" in Altium will not show this function
    var SchDoc;
    if (SchServer == null)
        return 1; //Incorrect file type
    SchDoc = SchServer.GetCurrentSchDocument();
    if (SchDoc == null) { //run this script in a wrong DOC
        showmessage("This is not a schematic document.");
        return 1; //Incorrect file type
    }
    return 0; //correct file type(sch)
}
//Create a Net label at (x,y) named <name>
function createNetLabel(x, y, rotate, name) {
    var SchDoc;
    var SchNetlabel;

    if (SchServer == null)
        return;
    SchDoc = SchServer.GetCurrentSchDocument();
    //Object setup
    SchNetlabel = SchServer.SchObjectFactory(eNetlabel, eCreate_GlobalCopy);
    SchNetlabel.Text = name;
    SchNetlabel.MoveToXY(MilsToCoord(x), MilsToCoord(y));
    SchNetlabel.RotateBy90(Point(MilsToCoord(x), MilsToCoord(y)), rotate);
    SchNetlabel.SetState_xSizeySize;
    SchNetlabel.GraphicallyInvalidate();
    SchDoc.RegisterSchObjectInContainer(SchNetlabel);
    //ShowMessage('create netLabel OK');
}

//Create a wire form(x_1, y_1) to (x_2, y_2)
function createWire(x_1, y_1, x_2, y_2) {
    var SchDoc;
    var SchWire;

    if (SchServer == null)
        return;
    SchDoc = SchServer.GetCurrentSchDocument();
    //Object setup
    SchWire = SchServer.SchObjectFactory(eWire, eCreate_GlobalCopy);
    SchWire.Location = Point(MilsToCoord(x_1), MilsToCoord(y_1));
    SchWire.InsertVertex = 1;
    SchWire.SetState_Vertex(1, Point(MilsToCOord(x_1), MilsToCoord(y_1)));
    SchWire.InsertVertex = 1;
    SchWire.SetState_Vertex(1, Point(MilsToCOord(x_2), MilsToCoord(y_2)));
    SchDoc.RegisterSchObjectInContainer(SchWire);
    //ShowMessage('create wire OK');
}
/* End of library */
