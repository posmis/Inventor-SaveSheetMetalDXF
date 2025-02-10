using Inventor;
using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace InventorAddin
{
    [Guid("12345678-ABCD-1234-ABCD-123456789ABC")]
    public class InventorAddin : ApplicationAddInServer
    {
        private Inventor.Application _inventorApp;
        private ButtonDefinition _button;
        private RibbonPanel _panel; // ссылка на панель для возможного удаления

        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            _inventorApp = addInSiteObject.Application;

            // Создаём кнопку
            CommandManager cmdMgr = _inventorApp.CommandManager;
            ControlDefinitions ctrlDefs = cmdMgr.ControlDefinitions;
            _button = ctrlDefs.AddButtonDefinition(
                "Save SheetMetal DXF",          // Текст на кнопке
                "MyInventorButton",        // Уникальный идентификатор
                CommandTypesEnum.kNonShapeEditCmdType,
                "{12345678-ABCD-1234-ABCD-123456789ABC}", // ID аддона
                "Save SheetMetal DXF", // Описание
                "Save SheetMetal DXF" // Tooltip
            );

            _button.OnExecute += ButtonClickHandler;

            // Получаем интерфейс пользователя и вкладку "Инструменты" в режиме детали (Part)
            UserInterfaceManager uiMgr = _inventorApp.UserInterfaceManager;
            Ribbon ribbon = uiMgr.Ribbons["Assembly"];
            RibbonTab toolsTab = ribbon.RibbonTabs["id_TabTools"];

            // Пытаемся получить существующую панель "Пользовательские команды"
            try
            {
                _panel = toolsTab.RibbonPanels["id_PanelUserCommands"];
            }
            catch
            {
                // Если такой панели нет, создаём новую
                _panel = toolsTab.RibbonPanels.Add("Пользовательские команды", "MyCommandsPanel", "{12345678-ABCD-1234-ABCD-123456789ABC}", "", true);
            }

            // Добавляем кнопку в выбранную панель
            _panel.CommandControls.AddButton(_button, false);
        }

        public void Deactivate()
        {
            // Опционально: удаляем созданную кнопку из панели
            if (_panel != null && _button != null)
            {
                try
                {
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

        private void ButtonClickHandler(NameValueMap context)
        {
            // Проверяем, что активен документ типа "Сборка"
            AssemblyDocument asmDoc = _inventorApp.ActiveDocument as AssemblyDocument;
            if (asmDoc == null)
            {
                MessageBox.Show("Откройте или создайте сборку!");
                return;
            }

            AssemblyComponentDefinition asmDef = asmDoc.ComponentDefinition;

            // Перебираем все компоненты сборки
            HashSet<string> uniqueComponentNames = new HashSet<string>(); // Хранит уникальные названия компонентов

            foreach (ComponentOccurrence occ in asmDef.Occurrences)
            {
                string componentName = occ.Definition.Document.DisplayName; // Берём оригинальное название файла

                // Добавляем только уникальные названия
                if (!uniqueComponentNames.Contains(componentName))
                {
                    uniqueComponentNames.Add(componentName);

                    // Определяем тип компонента
                    if (occ.Definition is PartComponentDefinition partDef)
                    {
                        if (partDef is SheetMetalComponentDefinition sheetMetalDef)
                        {
                            // Сохранение развертки в DXF
                            SaveSheetMetalDXF(sheetMetalDef, componentName);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Сохранение развёртки листовой детали в DXF формате
        /// </summary>
        private void SaveSheetMetalDXF(SheetMetalComponentDefinition sheetMetalDef, string componentName)
        {
            try
            {
                // Проверяем, есть ли развёртка
                if (sheetMetalDef.FlatPattern == null)
                {
                    MessageBox.Show($"Развертка отсутствует для {componentName}");
                    return;
                }



                // Указываем путь сохранения
                string dxfPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop), $"{componentName}.dxf");
                // Создаём папку DXF в текущей директории, если она не существует
                //string dxfFolderPath = System.IO.Path.Combine(System.Environment.CurrentDirectory, "DXF");
                //try
                //{
                //    if (!System.IO.Directory.Exists(dxfFolderPath))
                //    {
                //        System.IO.Directory.CreateDirectory(dxfFolderPath);
                //    }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show($"Ошибка при создании папки DXF:\n{ex.Message}");
                //    return; // Выход из метода в случае ошибки
                //}

                //// Указываем путь сохранения в папке DXF
                //string dxfPath = System.IO.Path.Combine(dxfFolderPath, $"{componentName}.dxf");


                // Создаём настройки экспорта
                TranslatorAddIn dxfTranslator = _inventorApp.ApplicationAddIns.ItemById["{C24E3AC4-122E-11D5-8E91-0010B541CD80}"] as TranslatorAddIn;
                if (dxfTranslator == null)
                {
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
            catch (Exception ex)
            {
                // MessageBox.Show($"Ошибка при сохранении DXF для {componentName}:\n{ex.Message}");
            }
        }


        public dynamic Automation => throw new NotImplementedException();
    }
}