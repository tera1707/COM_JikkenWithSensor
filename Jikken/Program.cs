using System.Net.Http.Headers;
using Vanara.PInvoke;
using static Vanara.PInvoke.SensorsApi;

namespace Jikken
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var mgr = new ISensorManager();
            foreach (var s in mgr.GetSensorsByCategory(Sensors.SENSOR_TYPE_AMBIENT_LIGHT).Enumerate())
            {
                Console.WriteLine($"{s.GetFriendlyName()}: {s.GetID()}");
                var vals = s.GetProperties(null);//?? throw new Exception();


                if (vals is not null)
                {
                    foreach (var pv in vals.Enumerate())
                        Console.WriteLine($"  {pv.Item2}");
                }
            }

            Console.ReadLine();
        }
    }
}