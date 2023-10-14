using System.Runtime.InteropServices;
using static ISensorManagerCs.Program;

namespace ISensorManagerCs
{
    internal class Program
    {
        const string SENSOR_TYPE_ACCELEROMETER_1D = "C04D2387-7340-4CC2-991E-3B18CB8EF2F4";
        const string SENSOR_TYPE_AMBIENT_LIGHT = "97F115C8-599A-4153-8894-D2D12899918A";
        const string SENSOR_DATA_TYPE_LIGHT_LEVEL_LUX = "E4C77CE2-DCB7-46E9-8439-4FEC548833A6";
#if false
        /// <summary>CLSID_Sensor</summary>
        [ComImport, Guid("E97CED00-523A-4133-BF6F-D3A2DAE7F6BA"), ClassInterface(ClassInterfaceType.None)]
        public class Sensor { };

        /// <summary>CLSID_SensorCollection</summary>
        [ComImport, Guid("79C43ADB-A429-469F-AA39-2F2B74B75937"), ClassInterface(ClassInterfaceType.None)]
        public class SensorCollection { };

        /// <summary>CLSID_SensorDataReport</summary>
        [ComImport, Guid("4EA9D6EF-694B-4218-8816-CCDA8DA74BBA"), ClassInterface(ClassInterfaceType.None)]
        public class SensorDataReport { };

        /// <summary>CLSID_SensorManager</summary>
        [ComImport, Guid("77A1C827-FCD2-4689-8915-9D613CC5FA3E")]//, ClassInterface(ClassInterfaceType.None)]
        public class SensorManager { };

        /// <summary>
        /// <para>Represents a collection of sensors, such as all the sensors connected to a computer.</para>
        /// <para>
        /// Retrieve a pointer to <c>ISensorCollection</c> by calling methods of the ISensorManager interface. In addition to the methods
        /// inherited from <c>IUnknown</c>, the <c>ISensorCollection</c> interface exposes the following methods.
        /// </para>
        /// </summary>
        // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nn-sensorsapi-isensorcollection
        [ComImport, Guid("23571E11-E545-4DD8-A337-B89BF44B10DF"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]//, CoClass(typeof(SensorCollection))]
        public interface ISensorCollection
        {
            /// <summary>Retrieves the sensor at the specified index in the collection.</summary>
            /// <param name="ulIndex"><c>ULONG</c> containing the index of the sensor to retrieve.</param>
            /// <returns>Address of an ISensor pointer that receives the pointer to the specified sensor.</returns>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensorcollection-getat HRESULT GetAt( [in] ULONG
            // ulIndex, [out] ISensor **ppSensor );
            ISensor GetAt(uint ulIndex);

            /// <summary>Retrieves the count of sensors in the collection.</summary>
            /// <returns>Address of a <c>ULONG</c> that receives the count.</returns>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensorcollection-getcount HRESULT GetCount( [out]
            // ULONG *pCount );
            uint GetCount();

            /// <summary>Adds a sensor to the collection.</summary>
            /// <param name="pSensor">Pointer to the ISensor interface for the sensor to add to the collection.</param>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensorcollection-add HRESULT Add( [in] ISensor
            // *pSensor );
            void Add(ISensor pSensor);

            /// <summary>Removes a sensor from the collection. The sensor is specified by a pointer to the ISensor interface to be removed.</summary>
            /// <param name="pSensor">Pointer to the ISensor interface to remove from the collection.</param>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensorcollection-remove HRESULT Remove( [in] ISensor
            // *pSensor );
            void Remove(ISensor pSensor);

            /// <summary>Removes a sensor from the collection. The sensor to be removed is specified by its ID.</summary>
            /// <param name="sensorID">The <c>GUID</c> of the sensor to remove from the collection.</param>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensorcollection-removebyid HRESULT RemoveByID( [in]
            // REFSENSOR_ID sensorID );
            void RemoveByID(in Guid sensorID);

            /// <summary>Empties the sensor collection.</summary>
            /// <remarks>
            /// This method calls <c>Release</c> on all sensor interface pointers in the collection and frees any memory used by the collection.
            /// </remarks>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensorcollection-clear HRESULT Clear();
            void Clear();
        }

        /// <summary>Provides methods for discovering and retrieving available sensors and a method to request sensor manager events.</summary>
        /// <remarks>
        /// <para>
        /// You retrieve a pointer to this interface by calling the COM <c>CoCreateInstance</c> method. If group policy does not allow creation
        /// of this object, <c>CoCreateInstance</c> will return <c>HRESULT_FROM_WIN32 (ERROR_ACCESS_DISABLED_BY_POLICY)</c>.
        /// </para>
        /// <para>Examples</para>
        /// <para>The following example code creates an instance of the sensor manager.</para>
        /// <para>
        /// <code>// Create the sensor manager. hr = CoCreateInstance(CLSID_SensorManager, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&amp;pSensorManager)); if(hr == HRESULT_FROM_WIN32(ERROR_ACCESS_DISABLED_BY_POLICY)) { // Unable to retrieve sensor manager due to // group policy settings. Alert the user. }</code>
        /// </para>
        /// </remarks>
        // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nn-sensorsapi-isensormanager
        [ComImport, Guid("BD77DB67-45A8-42DC-8D00-6DCF15F8377A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), CoClass(typeof(SensorManager))]
        public interface ISensorManager
        {
            /// <summary>Retrieves a collection containing all sensors associated with the specified category.</summary>
            /// <param name="sensorCategory">ID of the sensor category to retrieve.</param>
            /// <returns>Address of an ISensorCollection interface pointer that receives a pointer to the sensor collection requested.</returns>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensormanager-getsensorsbycategory HRESULT
            // GetSensorsByCategory( [in] REFSENSOR_CATEGORY_ID sensorCategory, [out] ISensorCollection **ppSensorsFound );
            public ISensorCollection GetSensorsByCategory(in Guid sensorCategory);

            /// <summary>Retrieves a collection containing all sensors associated with the specified type.</summary>
            /// <param name="sensorType">ID of the type of sensors to retrieve.</param>
            /// <returns>Address of an ISensorCollection interface pointer that receives the pointer to the sensor collection requested.</returns>
            // https://docs.microsoft.com/en-us/windows/win32/api/sensorsapi/nf-sensorsapi-isensormanager-getsensorsbytype HRESULT
            // GetSensorsByType( [in] REFSENSOR_TYPE_ID sensorType, [out] ISensorCollection **ppSensorsFound );
            public ISensorCollection GetSensorsByType(in Guid sensorType);

        }
        [ComImport, Guid("5FA08F80-2657-458E-AF75-46F73FA6AC5C"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown), CoClass(typeof(Sensor))]
        public interface ISensor
        {

        }
#endif


        static void Main(string[] args)
        {
#if false
            var sensorManager = (ISensorManager)Activator.CreateInstance(CLSID.SensorManagerType)!;
#else
            var sensorManager = new SensorManager();
#endif
            if (sensorManager is ISensorManager sm)
            {
                //sm.GetSensorsByCategory(new Guid(SENSOR_TYPE_AMBIENT_LIGHT), out var sc);
                var sc = sm.GetSensorsByCategory(new Guid(SENSOR_TYPE_AMBIENT_LIGHT));

                if (sc is ISensorCollection c)
                {
                    //var count = c.GetCount(out var count);
                    var count = c.GetCount();
                    for (uint i = 0; i < count; i++)
                    {
                        //var sensor = sc.GetAt(i, out var sensor);
                        var sensor = sc.GetAt(i);

                        sensor.GetData(out var report);
                        report.GetSensorValue(new PROPERTYKEY(new Guid(SENSOR_DATA_TYPE_LIGHT_LEVEL_LUX), 2), out var x);

                        Console.WriteLine($"明るさは {x.fltVal} ルクスです");
                    }
                }
            }
            Console.ReadLine();
        }

#if false
        public static class CLSID
        {
            public static readonly Guid SensorManager = new Guid("77A1C827-FCD2-4689-8915-9D613CC5FA3E");
            public static readonly Type SensorManagerType = Type.GetTypeFromCLSID(SensorManager)!;
        }
#else
        // これがいわゆる「CoClass」だと思われる（たぶん）
        [ComImport]
        [Guid("77A1C827-FCD2-4689-8915-9D613CC5FA3E")]
        internal class SensorManager
        {
        }
#endif
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("BD77DB67-45A8-42DC-8D00-6DCF15F8377A")]
        internal interface ISensorManager
        {
            //[PreserveSig]
            //public long GetSensorsByCategory([In] ref Guid sensorCategory, [Out] out ISensorCollection ppSensorsFound);
            public ISensorCollection GetSensorsByCategory(in Guid sensorCategory);

            // 以下略（使わないものは省略してOK）
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("23571E11-E545-4DD8-A337-B89BF44B10DF")]
        internal interface ISensorCollection
        {
            //[PreserveSig]
            //public int GetAt([In] uint ulIndex, [Out] out ISensor ppSensor);
            public ISensor GetAt([In] uint ulIndex);
            //[PreserveSig]
            //public int GetCount([Out] out uint count);
            public uint GetCount();


        }

        [ComImport]
        [Guid("79C43ADB-A429-469F-AA39-2F2B74B75937")]
        public class SensorCollection
        {
        }
        //



        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("5FA08F80-2657-458E-AF75-46F73FA6AC5C")]
        internal interface ISensor
        {
            [PreserveSig]
            public long GetID(ulong ulIndex, out Guid pID);
            [PreserveSig]
            public void dummy2();
            [PreserveSig]
            public void dummy3();
            [PreserveSig]
            public void dummy4();
            [PreserveSig]
            public void dummy5();
            [PreserveSig]
            public void dummy6();
            [PreserveSig]
            public void dummy7();
            [PreserveSig]
            public void dummy8();
            [PreserveSig]
            public void dummy9();
            [PreserveSig]
            public void dummy10();
            [PreserveSig]
            public void GetData([Out] out ISensorDataReport ppDataReport);
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("0AB9DF9B-C4B5-4796-8898-0470706A2E1D")]
        internal interface ISensorDataReport
        {
            [PreserveSig]
            public void dummy1();
            [PreserveSig]
            public long GetSensorValue([In] PROPERTYKEY pKey, [Out] out PROPVARIANT pValue);
        }


        [StructLayout(LayoutKind.Explicit)]
        public struct PROPVARIANT
        {
            [FieldOffset(0)]
            public ushort vt;

            [FieldOffset(2)]
            public ushort wReserved1;

            [FieldOffset(4)]
            public ushort wReserved2;

            [FieldOffset(6)]
            public ushort wReserved3;

            [FieldOffset(8)]
            public IntPtr pwszVal;

            [FieldOffset(8)]
            public float fltVal;
        }



        [StructLayout(LayoutKind.Sequential, Pack =4)]
        public struct PROPERTYKEY
        {
            public Guid fmtid;
            public uint pid;

            public PROPERTYKEY(Guid key, uint id)
            {
                fmtid = key;
                pid = id;
            }
        }
    }
}



// 明るさセンサーの情報
//https://learn.microsoft.com/en-us/windows/win32/sensorsapi/sensor-category-light

// PROPVARIANT
//https://learn.microsoft.com/ja-jp/windows/win32/api/propidl/ns-propidl-propvariant

// PROPVARIANTの変換？
// https://blog.shibayan.jp/entry/20220504/1651658855




