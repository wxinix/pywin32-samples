using System;
using System.Runtime.InteropServices;
using System.Xml.Linq;

namespace PythonCOMTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            dynamic comObj = null;
            dynamic personObj = null;

            try
            {
                Type comType = Type.GetTypeFromProgID("Python.MyCOMObject");
                comObj = Activator.CreateInstance(comType);

                // Call methods on the COM object
                personObj = comObj.get_person();

                // Access the name and age properties directly
                string name = personObj.name;
                int age = personObj.age;

                Console.WriteLine("Name: " + name);
                Console.WriteLine("Age: " + age);
            }
            catch(COMException ex)
            {
                // Handle any COM exception that might occur
                Console.WriteLine("Error: " + ex.Message);
            }
            finally
            {
                // Release the COM object
                if(comObj != null)
                {
                    Marshal.ReleaseComObject(comObj);
                }

                // Release the COM object
                if(personObj != null)
                {
                    Marshal.ReleaseComObject(personObj);
                }
            }
        }
    }
}
