#include "./ExcelToFlatBuffer.h"

#include "GameLib.h"

#include <oleidl.h>
#include <shellapi.h>

using DropCallback = std::function<void(const wchar_t* xlsxPath)>;

/**
 * 拖拽管理器类
 */
class DropManager : public IDropTarget {
public:
    DropManager(DropCallback callback) : m_callback(callback) {}

    ULONG AddRef() { return 1; }
    ULONG Release() { return 0; }

    HRESULT QueryInterface(REFIID riid, void** ppvObject) {
        if (riid == IID_IDropTarget) {
            *ppvObject = this;
            return S_OK;
        }
        *ppvObject = NULL;
        return E_NOINTERFACE;
    }

    HRESULT DragEnter(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
        *pdwEffect &= DROPEFFECT_COPY;
        return S_OK;
    }

    HRESULT DragLeave() {
        return S_OK;
    }

    HRESULT DragOver(DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
        *pdwEffect &= DROPEFFECT_COPY;
        return S_OK;
    }

    HRESULT Drop(IDataObject* pDataObj, DWORD grfKeyState, POINTL pt, DWORD* pdwEffect) {
        *pdwEffect &= DROPEFFECT_COPY;

        FORMATETC fmtetc = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
        STGMEDIUM stgmedium;
        ZeroMemory(&stgmedium, sizeof(stgmedium));

        HRESULT hr = pDataObj->GetData(&fmtetc, &stgmedium);
        if (SUCCEEDED(hr)) {
            HDROP hDrop = (HDROP)stgmedium.hGlobal;
            wchar_t xlsxPath[MAX_PATH] = { 0 };
            DragQueryFileW(hDrop, 0, xlsxPath, MAX_PATH);
            m_callback(xlsxPath);
            DragFinish(hDrop);
            ReleaseStgMedium(&stgmedium);
        }

        return S_OK;
    }

private:
    DropCallback m_callback;
};

class AutoOleDragDrop {
public:
    AutoOleDragDrop(const wchar_t* className, DropCallback callback) : m_hwnd(NULL), m_pDropTarget(NULL), m_ok(false) {
        m_hwnd = FindWindowW(className, NULL);
        if (!m_hwnd) return;

        if (FAILED(OleInitialize(NULL))) return;

        m_pDropTarget = new DropManager(callback);
        if (FAILED(RegisterDragDrop(m_hwnd, m_pDropTarget))) {
            delete m_pDropTarget;
            OleUninitialize();
            return;
        }

        m_ok = true;
    }

    ~AutoOleDragDrop() {
        if (m_ok) {
            if (m_hwnd) RevokeDragDrop(m_hwnd);
            if (m_pDropTarget) m_pDropTarget->Release();
            OleUninitialize();
        }
    }

    bool IsOk() const { return m_ok; }

private:
    HWND m_hwnd;
    IDropTarget* m_pDropTarget;
    bool m_ok;
};

#define WIDTH 640
#define HEIGHT 480

int main333()
{
    GameLib game;
    game.Open(WIDTH, HEIGHT, "Tools", true);

    //ShowWindow(GetConsoleWindow(), SW_HIDE);

    AutoOleDragDrop oleDragDrop(L"GameLibWindowClass", [](const wchar_t* xlsxPath) {
        std::string excelFile = UTF8ToGB2312(WcharToChar(xlsxPath).c_str());
        std::string outputFile = MakeDesPath(excelFile, ".bytes");
        ExcelToFlatBuffer converter;
        // 执行转换
        if (!converter.Convert(excelFile, outputFile)) {
            std::cerr << "转换失败: " << converter.GetLastError() << std::endl;
        }
        });
    if (!oleDragDrop.IsOk()) {
        game.ShowMessage("Failed to initialize drag & drop", "Error", MESSAGEBOX_OK);
    }

    int x = 320, y = 240;

    // 定义按钮位置和大小
    int btnW = 100;
    int btnH = 60;
    int btnX = 0;
    int btnY = HEIGHT - btnH;
    // 定义按钮文本
    const char* btnText = "Help!";
    // 定义消息框显示的文本
    const char* message = "将Excel文件拖入窗口可生成配置！";

    while (!game.IsClosed()) {
        if (game.IsKeyDown(KEY_LEFT))  x -= 3;
        if (game.IsKeyDown(KEY_RIGHT)) x += 3;
        if (game.IsKeyDown(KEY_UP))    y -= 3;
        if (game.IsKeyDown(KEY_DOWN))  y += 3;

        game.Clear(COLOR_BLACK);
        // 绘制按钮，并检查是否被点击
        if (game.Button(btnX, btnY, btnW, btnH, btnText, COLOR_BLUE)) {
            
            // 按钮被点击时显示消息框
            game.ShowMessage(GB2312ToUTF8(message), "Button Clicked", MESSAGEBOX_OK);
        }

        game.FillCircle(x, y, 15, COLOR_CYAN);
        game.DrawText(10, 10, "Up/Down/Left/Right to move!", COLOR_WHITE);
        game.Update();
        game.WaitFrame(60);
    }

    return 0;
}
