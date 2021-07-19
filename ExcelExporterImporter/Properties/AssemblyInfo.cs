using System.Reflection;
using System.Runtime.InteropServices;
using log4net.Config;

//#error Change the assembly information and set new GUID then comment this line
// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.

[assembly: AssemblyConfiguration("BIMO")]
[assembly: AssemblyTrademark("BIMOne")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.

[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM

[assembly: Guid("5C4B1590-780E-4732-A080-732AE5AF249F")]
[assembly: XmlConfigurator(Watch = true)]