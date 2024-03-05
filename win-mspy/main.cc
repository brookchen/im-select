#define UNICODE

#include <Windows.h>
#include <UIAutomation.h>
#include <iostream>
#include <string>
#include <comip.h>
#include <comdef.h>
#include <vector>
#include <regex>
#include <WinUser.h>
#include <cctype>
#include <algorithm>
#include <locale>
#include <sstream>

using namespace std;

// command line options
struct CliOptions
{
    // no prefix
    wstring mode;
    // -k= 
    wstring switch_keys;
    // -t=
    wstring taskbar_name;
    // -i=
    wstring ime_capture_re;
    wregex ime_capture;
    //-l=
    vector<wstring> ime_vec;
};

wstring get_element_name(IUIAutomationElement* pElement)
{
    BSTR name;
    HRESULT hr = pElement->get_CurrentName(&name);
    if (FAILED(hr))
    {
        return L"";
    }
    wstring res(name, SysStringLen(name));
    SysFreeString(name);
    return res;
}

vector<wstring> split_string(const wstring& str, const wstring& delim)
{
    wregex re(delim);
    wsregex_token_iterator first{ str.begin(), str.end(), re, -1 }, last;
    return { first, last };
}

SHORT vk_from_text(const wstring& text) {
    if (text == L"shift") {
        return VK_SHIFT;
    }
    if (text == L"ctrl") {
        return VK_CONTROL;
    }
    if (text == L"alt") {
        return VK_MENU;
    }
    if (text == L"win") {
        return VK_LWIN;
    }
    if (text == L"space") {
        return VK_SPACE;
    }
    return 0;
}

vector<INPUT> get_input_from_string(wstring str)
{
    vector<INPUT> result;
    transform(str.begin(), str.end(), str.begin(), ::tolower);
    auto keys = split_string(str, L"\\+");
    transform(keys.begin(), keys.end(), back_insert_iterator(result), [](const wstring& key) {
        INPUT input = { 0 };
        input.type = INPUT_KEYBOARD;
        input.ki.wVk = vk_from_text(key);
        return input;
        });
    transform(keys.rbegin(), keys.rend(), back_insert_iterator(result), [](const wstring& key) {
        INPUT input = { 0 };
        input.type = INPUT_KEYBOARD;
        input.ki.wVk = vk_from_text(key);
        input.ki.dwFlags = KEYEVENTF_KEYUP;
        return input;
        });
    return result;
}

// default chinese options
CliOptions chinese_options()
{
    CliOptions options;
    options.taskbar_name = L"任务栏";
    options.ime_capture_re = L"托盘输入指示器\\s+(\\w+)";
    options.switch_keys = L"shift";
    return options;
}

void split_str(const wstring& str, wchar_t delim, vector<wstring>& vec) {
    wstringstream sstream(str);
    wstring item;
    while (getline(sstream, item, delim)) {
        vec.push_back(item);
    }
}

// parse command line options
CliOptions parse_options(int argc, wchar_t* argv[])
{
    CliOptions options = chinese_options();
    for (int i = 1; i < argc; i++)
    {
        auto arg = argv[i];
        if (arg[0] == L'-')
        {
            auto pos = wcschr(arg, L'=');
            if (pos)
            {
                auto key = wstring(arg + 1, pos);
                auto value = wstring(pos + 1);
                if (key == L"k")
                {
                    options.switch_keys = value;
                }
                else if (key == L"t")
                {
                    options.taskbar_name = value;
                }
                else if (key == L"i")
                {
                    options.ime_capture_re = value;
                }
                else if (key == L"l")
                {
                    split_str(value, ',', options.ime_vec);
                }
            }
        }
        else
        {
            options.mode = arg;
        }
    }
    if (options.ime_capture_re.length() > 0)
    {
        options.ime_capture = wregex(options.ime_capture_re);
    }
    if (options.ime_vec.size() > 0)
    {
        try
        {
            int mode = stoi(options.mode);
            if (mode >= 0 && mode < options.ime_vec.size())
            {
                options.mode = options.ime_vec[mode];
            }
        }
        catch (exception&)
        {
            //do nothing
        }
    }
    return options;
}

void print_options(const CliOptions& options)
{
    wcout << L"taskbar name: " << options.taskbar_name << endl;
    wcout << L"ime capture: " << options.ime_capture_re << endl;
    wcout << L"switch keys: " << options.switch_keys << endl;
    wcout << L"mode: " << options.mode << endl;
    wcout << L"ime lists: ";
    for (auto& ime : options.ime_vec)
    {
        wcout << ime << " ";
    }
    wcout << endl;
}

#define IF_FAIL_GOTO(ptr, label) if (FAILED(ptr)){goto label;}
#define RELEASE_PTR(ptr) if (ptr != nullptr) {ptr->Release(); ptr=nullptr;}
int get_ime_mode(const CliOptions& options, wstring& mode)
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr))
    {
        wprintf(L"CoInitialize failed, HR:0x%08x\n", hr);
        return 1;
    }
    int ret = 1;
    IUIAutomation* automation = nullptr;
    IUIAutomationElement* desktop = nullptr;
    IUIAutomationElement* task_bar = nullptr;
    IUIAutomationCondition* condition1 = nullptr;
    IUIAutomationCondition* condition2 = nullptr;
    IUIAutomationElementArray* buttons_arr = nullptr;

    hr = CoCreateInstance(__uuidof(CUIAutomation8), NULL,
        CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&automation));
    IF_FAIL_GOTO(hr, failed);
    hr = automation->GetRootElement(&desktop);
    IF_FAIL_GOTO(hr, failed);
    hr = automation->CreatePropertyCondition(UIA_NamePropertyId, _variant_t(options.taskbar_name.c_str()), &condition1);
    IF_FAIL_GOTO(hr, failed);
    hr = desktop->FindFirst(TreeScope_Children, condition1, &task_bar);
    IF_FAIL_GOTO(hr, failed);
    hr = automation->CreatePropertyCondition(UIA_ControlTypePropertyId, _variant_t(UIA_ButtonControlTypeId), &condition2);
    IF_FAIL_GOTO(hr, failed);
    hr = task_bar->FindAll(TreeScope_Descendants, condition2, &buttons_arr);
    IF_FAIL_GOTO(hr, failed);

    int length = 0;
    buttons_arr->get_Length(&length);
    for (int i = 0; i < length; i++)
    {
        IUIAutomationElement* button;
        hr = buttons_arr->GetElement(i, &button);
        if (SUCCEEDED(hr))
        {
            auto name = get_element_name(button);
            button->Release();
            wsmatch match;
            if (regex_search(name, match, options.ime_capture))
            {
                mode = match[1];
                ret = 0;
                break;
            }
        }
    }
failed:
    RELEASE_PTR(buttons_arr);
    RELEASE_PTR(condition2);
    RELEASE_PTR(task_bar);
    RELEASE_PTR(condition1);
    RELEASE_PTR(desktop);
    RELEASE_PTR(automation);
    CoUninitialize();
    return ret;
}

int _cdecl wmain(int argc, WCHAR* argv[])
{
    ios::sync_with_stdio(false);
    locale::global(std::locale(""));

    auto options = parse_options(argc, argv);
    //print_options(options);

    wstring cur_mode;
    int res = get_ime_mode(options, cur_mode);
    if (res != 0)
    {
        return 1;
    }

    if (options.mode.empty())      // output current mode
    {
        if (options.ime_vec.size() > 0)
        {
            for (int i = 0; i < options.ime_vec.size(); i++)
            {
                auto& ime = options.ime_vec[i];
                if (ime == cur_mode)
                {
                    wcout << i << endl;
                    break;
                }
            }
        }
        else
        {
            wcout << cur_mode << endl;
        }
    }
    else     // do switch
    {
        //wcout << L"current ime:" << ime_button.current_mode << endl;
        if (options.mode != cur_mode)
        {
            auto input = get_input_from_string(options.switch_keys);
            SendInput((UINT)input.size(), input.data(), sizeof(input[0]));
        }
    }
    return 0;
}
