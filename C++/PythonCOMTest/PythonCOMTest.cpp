#include <iostream>
#include <Windows.h>
#include <atlbase.h>
#include <atlcom.h>
#include <string>
#include <stdexcept>

CComVariant GetProperty(IDispatch* pObject, BSTR propName) {
  DISPID dispid;
  HRESULT hr = pObject->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
  if (FAILED(hr)) {
    throw std::runtime_error("Failed to get property ID");
  }

  DISPPARAMS params = { nullptr, nullptr, 0, 0 };
  CComVariant result;
  hr = pObject->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_PROPERTYGET, &params, &result, nullptr, nullptr);
  if (FAILED(hr)) {
    throw std::runtime_error("Failed to get property value");
  }

  return result;
}

int main() {
  HRESULT hr = CoInitialize(NULL);

  if (SUCCEEDED(hr)) {
    CLSID clsid;
    hr = CLSIDFromProgID(L"Python.MyCOMObject", &clsid);

    if (SUCCEEDED(hr)) {
      CComPtr<IClassFactory> classFactory;
      hr = CoGetClassObject(clsid, CLSCTX_INPROC_SERVER, NULL, IID_PPV_ARGS(&classFactory));

      if (SUCCEEDED(hr)) {
        CComPtr<IDispatch> comInstance;
        hr = classFactory->CreateInstance(NULL, IID_PPV_ARGS(&comInstance));

        if (SUCCEEDED(hr)) {
          BSTR methodName = SysAllocString(L"get_person");
          if (methodName != nullptr) {
            DISPID dispid;
            hr = comInstance->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_SYSTEM_DEFAULT, &dispid);
            SysFreeString(methodName);

            if (SUCCEEDED(hr)) {
              DISPPARAMS params = { nullptr, nullptr, 0, 0 };
              CComVariant result;
              hr = comInstance->Invoke(dispid, IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &params, &result, nullptr, nullptr);

              if (SUCCEEDED(hr) && result.vt == VT_DISPATCH) {
                CComPtr<IDispatch> person = result.pdispVal;

                try {
                  BSTR propName = SysAllocString(L"name");
                  if (propName != nullptr) {
                    CComVariant name = GetProperty(person, propName);
                    SysFreeString(propName);

                    propName = SysAllocString(L"age");
                    if (propName != nullptr) {
                      CComVariant age = GetProperty(person, propName);
                      SysFreeString(propName);

                      std::wcout << L"Person's Name: " << name.bstrVal << std::endl;
                      std::wcout << L"Person's Age: " << age.intVal << std::endl;
                    }
                  }
                }
                catch (const std::runtime_error& e) {
                  std::cerr << e.what() << std::endl;
                }
              }
              else {
                std::cerr << "Failed to call the 'get_person' method or return value is not a dispatch object." << std::endl;
              }
            }
            else {
              std::cerr << "Failed to get the 'get_person' method ID." << std::endl;
            }
          }
        }
        else {
          std::cerr << "Failed to create an instance of the COM server." << std::endl;
        }
      }
      else {
        std::cerr << "Failed to get the class object of the COM server." << std::endl;
      }
    }
    else {
      std::cerr << "Failed to get the CLSID for the COM server." << std::endl;
    }
    CoUninitialize();
  }
  else {
    std::cerr << "Failed to initialize COM library." << std::endl;
  }

  return 0;
}
