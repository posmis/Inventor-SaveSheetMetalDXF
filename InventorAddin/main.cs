using Inventor;

using System;
using System.IO;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

using InventorAddin.Properties;
using System.Globalization;

namespace InventorAddin {
    [Guid("12345678-ABCD-1234-ABCD-123456789ABC")]

    public class InventorAddin : ApplicationAddInServer {
        private Inventor.Application _inventorApp;
        private ButtonDefinition _button;
        private RibbonPanel _panel; // ссылка на панель для возможного удаления

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime) {
            _inventorApp = addInSiteObject.Application;

            //System.Drawing.Bitmap smallIconBmp = new System.Drawing.Bitmap("C:\\InventorAddin.Small.bmp");
            //System.Drawing.Bitmap largeIconBmp = new System.Drawing.Bitmap("C:\\InventorAddin.Large.bmp");

            //object smallIconDisp = AxHostConverter.ImageToPictureDisp(smallIconBmp);
            //object largeIconDisp = AxHostConverter.ImageToPictureDisp(largeIconBmp);

            // Создаём кнопку
            CommandManager cmdMgr = _inventorApp.CommandManager;
            ControlDefinitions ctrlDefs = cmdMgr.ControlDefinitions;
            _button = ctrlDefs.AddButtonDefinition(
                "Save in DXF",          // Текст на кнопке
                "MyInventorButton",        // Уникальный идентификатор
                CommandTypesEnum.kNonShapeEditCmdType,
                "{12345678-ABCD-1234-ABCD-123456789ABC}", // ID аддона
                "Save SheetMetal DXF", // Описание
                "Save SheetMetal DXF"
                //smallIconDisp,
                //largeIconDisp
            );

            _button.OnExecute += ButtonClickHandler;

            // Получаем интерфейс пользователя и вкладку "Инструменты" в режиме детали (Part)
            UserInterfaceManager uiMgr = _inventorApp.UserInterfaceManager;
            Ribbon ribbon = uiMgr.Ribbons["Assembly"];
            RibbonTab toolsTab = ribbon.RibbonTabs["id_TabTools"];

            // Пытаемся получить существующую панель "Пользовательские команды"
            try {
                _panel = toolsTab.RibbonPanels["id_PanelUserCommands"];
            }
            catch {
                // Если такой панели нет, создаём новую
                _panel = toolsTab.RibbonPanels.Add("Пользовательские команды", "MyCommandsPanel", "{12345678-ABCD-1234-ABCD-123456789ABC}", "", true);
            }

            // Добавляем кнопку в выбранную панель
            _panel.CommandControls.AddButton(_button, false);
        }

        public void Deactivate() {
            // Опционально: удаляем созданную кнопку из панели
            if (_panel != null && _button != null) {
                try {
                    // Если панель была создана нами, можно её также удалить,
                    // но будьте осторожны – удаление панели может повлиять на UI.
                    _panel.CommandControls["MyInventorButton"].Delete();
                }
                catch { }
            }
            if (_button != null)
                Marshal.ReleaseComObject(_button);

            _button = null;
            _panel = null;
            _inventorApp = null;
        }

        public void ExecuteCommand(int commandID) { }

        private void ButtonClickHandler(NameValueMap context) {
            // Проверяем, что активен документ типа "Сборка"
            AssemblyDocument asmDoc = _inventorApp.ActiveDocument as AssemblyDocument;
            if (asmDoc == null) {
                MessageBox.Show("Откройте или создайте сборку!");
                return;
            }
            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;

            Dictionary<string, int> componentCounts = new Dictionary<string, int>();

            foreach (ComponentOccurrence occ in asmDef.Occurrences) {
                string componentName = occ.Definition.Document.DisplayName; // Берём оригинальное название файла

                // Подсчитываем количество вхождений
                if (componentCounts.ContainsKey(componentName)) {
                    componentCounts[componentName]++;
                } else {
                    componentCounts[componentName] = 1;
                }
            }

            // Перебираем все компоненты сборки
            HashSet<string> uniqueComponentNames = new HashSet<string>(); // Хранит уникальные названия компонентов

            foreach (var entry in componentCounts) {
                string componentName = entry.Key;
                int count = entry.Value;

                // Поиск первого вхождения компонента вручную
                ComponentOccurrence firstOccurrence = null;
                foreach (ComponentOccurrence occ in asmDef.Occurrences) {
                    if (occ.Definition.Document.DisplayName == componentName) {
                        firstOccurrence = occ;
                        break;
                    }
                }
                if (firstOccurrence != null && firstOccurrence.Definition is PartComponentDefinition partDef) {
                    if (partDef is SheetMetalComponentDefinition sheetMetalDef) {
                        // Сохранение развертки в DXF
                        SaveSheetMetalDXF(sheetMetalDef, componentName, count);
                    }
                }
            }
        }

        /// <summary>
        /// Сохранение развёртки листовой детали в DXF формате
        /// </summary>
        private void SaveSheetMetalDXF(SheetMetalComponentDefinition sheetMetalDef, string fileName, int fileCount) {
            try {
                // Проверяем, есть ли развёртка
                if (sheetMetalDef.FlatPattern == null) {
                    MessageBox.Show($"Развертка отсутствует для {fileName}");
                    return;
                }

                // Указываем путь сохранения
                Double thicknessParam = sheetMetalDef.Thickness.Value * 10;
                string thicknessParamStr = thicknessParam.ToString("F1", CultureInfo.InvariantCulture);
                string componentName = $"{System.IO.Path.GetFileNameWithoutExtension(fileName)}-t{thicknessParamStr}мм-{fileCount}шт";
                string dxfPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), $"{componentName}.dxf");

                // Создаём настройки экспорта
                TranslatorAddIn dxfTranslator = _inventorApp.ApplicationAddIns.ItemById["{C24E3AC4-122E-11D5-8E91-0010B541CD80}"] as TranslatorAddIn;
                if (dxfTranslator == null) {
                    MessageBox.Show("DXF-транслятор не найден или не поддерживает сохранение.");
                    return;
                }

                TranslationContext context = _inventorApp.TransientObjects.CreateTranslationContext();
                context.Type = IOMechanismEnum.kFileBrowseIOMechanism;

                NameValueMap options = _inventorApp.TransientObjects.CreateNameValueMap();
                options.Value["FlatPatternView"] = true;

                DataMedium dataMedium = _inventorApp.TransientObjects.CreateDataMedium();
                dataMedium.FileName = dxfPath;

                // Сохраняем развертку в DXF
                dxfTranslator.SaveCopyAs(sheetMetalDef.FlatPattern, context, options, dataMedium);
            }
            catch (Exception ex) {
                MessageBox.Show($"Ошибка при сохранении DXF для {fileName}:\n{ex.Message}");
            }
        }

        public dynamic Automation => throw new NotImplementedException();
    }
}