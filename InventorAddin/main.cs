using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;

using Inventor;
using InventorAddin.Properties;


namespace InventorAddin {
    [Guid("80960a53-04de-42b0-ab12-ba9059b555a1")]

    public class InventorAddin : ApplicationAddInServer {
        private Inventor.Application _inventorApp;
        private ButtonDefinition _button;
        private RibbonPanel _panel;

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime) {
            try {
                _inventorApp = addInSiteObject.Application;
                
                Inventor.IPictureDisp smallIcon = PictureDispConverter.GetIPictureDispFromImage(Resources.small_ico);
                Inventor.IPictureDisp largeIcon = PictureDispConverter.GetIPictureDispFromImage(Resources.large_ico);


                // Создаём кнопку
                CommandManager cmdMgr = _inventorApp.CommandManager;
                ControlDefinitions ctrlDefs = cmdMgr.ControlDefinitions;
                _button = ctrlDefs.AddButtonDefinition(
                    "SheetMetal to DXF",          // Текст на кнопке
                    "ExportSheetMetalDXF",        // Уникальный идентификатор
                    CommandTypesEnum.kNonShapeEditCmdType,
                    "{80960a53-04de-42b0-ab12-ba9059b555a1}", // ID аддона
                    "Export all sheet metal parts to DXF", // Описание
                    "Export sheet metal as DXF",
                    smallIcon, largeIcon
                );

                _button.OnExecute += ButtonClickHandler;

                // Получаем интерфейс пользователя и вкладку "Инструменты" в режиме сборки (Assembly)
                UserInterfaceManager uiMgr = _inventorApp.UserInterfaceManager;
                Ribbon ribbon = uiMgr.Ribbons["Assembly"];
                RibbonTab toolsTab = ribbon.RibbonTabs["id_TabTools"];

                // Пытаемся получить существующую панель "Пользовательские команды"
                try {
                    _panel = toolsTab.RibbonPanels["id_PanelUserCommands"];
                } catch {
                    // Если такой панели нет, создаём новую
                    _panel = toolsTab.RibbonPanels.Add("Пользовательские команды", "MyCommandsPanel", "{12345678-ABCD-1234-ABCD-123456789ABC}", "", true);
                }

                // Добавляем кнопку в выбранную панель
                _panel.CommandControls.AddButton(_button, false);
            } catch (Exception ex) {
                MessageBox.Show("Ошибка при активации аддона: " + ex.Message);
            }
        }

        public void Deactivate() {
            // Опционально: удаляем созданную кнопку из панели
            if (_panel != null && _button != null) {
                try {
                    // Если панель была создана нами, можно её также удалить,
                    // но будьте осторожны – удаление панели может повлиять на UI.
                    _button.OnExecute -= ButtonClickHandler;
                    _panel.CommandControls["ExportSheetMetalDXF"].Delete();
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

        public class PictureDispConverter : AxHost {
            private PictureDispConverter() : base(string.Empty) { }
            public static Inventor.IPictureDisp GetIPictureDispFromImage(System.Drawing.Icon icon) {
                return (Inventor.IPictureDisp)GetIPictureDispFromPicture(icon.ToBitmap());
            }
        }
    }
}