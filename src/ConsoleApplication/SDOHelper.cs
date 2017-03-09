using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SAGE
{
    public static class SDOHelper

    {

        /// <summary>

        /// Writes to an SDO field

        /// </summary>

        /// <param name="oObject">The object to write to</param>

        /// <param name="fname">Name of the required field</param>

        /// <param name="value">Value you wish to write</param>

        public static void Write(Object oObject, String fname, Object value)

        {

            //Stores the required field name in an object

            Object Field = fname;



            //Uses reflection to access the objects fields collection

            //then passes a reference to the field required, setting its value

            //to what is required

            ((SageDataObject220.Fields)oObject.GetType().InvokeMember

              ("Fields", System.Reflection.BindingFlags.GetProperty, null, oObject, null)).Item

              (ref Field).Value = value;

        }





        /// <summary>

        /// Reads and SDO field and returns its value as an object

        /// </summary>

        /// <param name="oObject">The object to read from</param>

        /// <param name="fname">Name of the required field</param>

        /// <returns>Returns an Object containing the value from the field</returns>

        public static Object Read(Object oObject, String fname)

        {

            //Stores the required field name in an object

            Object Field = fname;



            //Uses reflection to access the objects fields collection

            //then passes a reference to the field required, returning its value

            return ((SageDataObject220.Fields)oObject.GetType().InvokeMember

              ("Fields", System.Reflection.BindingFlags.GetProperty, null, oObject, null)).Item

              (ref Field).Value;

        }





        /// <summary>

        /// Invokes the Add() method of an items collection

        /// </summary>

        /// <param name="oObject">The items collection you wish to invoke the Add() method on</param>

        /// <returns></returns>

        public static Object Add(Object oObject)

        {

            //Uses reflection to invoke the Add() Method on the required object

            return oObject.GetType().InvokeMember

              ("Add", System.Reflection.BindingFlags.InvokeMethod, null, oObject, null);

        }

    }

}
