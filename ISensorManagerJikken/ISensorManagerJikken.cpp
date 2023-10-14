#include <stdio.h>
#include <InitGuid.h>
#include <SensorsApi.h>
#include <Sensors.h>
#include <string>

#pragma comment(lib, "Sensorsapi.lib")

// https://learn.microsoft.com/ja-jp/windows/win32/sensorsapi/sensor-category-light

int main()
{
    ISensorManager* pSensorManager;
    ISensorCollection* pMotionSensorCollection;
    ISensor* pAmbientSensor;

    CoInitialize(NULL);

    auto hr = CoCreateInstance(CLSID_SensorManager, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pSensorManager));

    pSensorManager->GetSensorsByCategory(SENSOR_TYPE_AMBIENT_LIGHT, &pMotionSensorCollection);

    ULONG count = 0;
    pMotionSensorCollection->GetCount(&count);

    for (int i = 0; i < count; i++)
    {
        pMotionSensorCollection->GetAt(i, &pAmbientSensor);

        ISensorDataReport* pData;
        pAmbientSensor->GetData(&pData);
        //PROPERTYKEY
        // 照度センサの値(ルクス)を取る
        PROPVARIANT x = {};
        pData->GetSensorValue(SENSOR_DATA_TYPE_LIGHT_LEVEL_LUX, &x);

        wprintf(std::to_wstring(x.fltVal).c_str());
    }

    CoUninitialize();
}
