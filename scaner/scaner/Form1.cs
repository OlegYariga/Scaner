using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.IO;
using WIA;
using System.Drawing.Drawing2D;
using System.Diagnostics;




namespace scaner
{


    public partial class Form1 : Form
    {
        public static int help = 0;
        public Form1()
        {
            try
            {
                InitializeComponent(); // инициализация компонентов
            }catch(Exception exep)// исключение при инициализации (может возникать, если нет нужных dll на ПК)
            {
                MessageBox.Show("На компьютере не обнаружено динамической библиотеки WIA. Ошибка: "+ exep.Message, "Критическая ошибка!");
                Process.GetCurrentProcess().Kill();//закрытие программы через процесс программы

            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                Icon icon = Icon.ExtractAssociatedIcon("scanner_104151.ico");
                this.Icon = icon;
            }catch(Exception exx)
            {
                MessageBox.Show("Не загружен один из файлов ресурсов. Работа программы продолжится\n\n Ошибка: возможно, не найден файл значка программы по адресу " + exx.Message, "Ошибка");
            }
            this.DoubleBuffered = true;// двойная буферизация
            pictureBox1.SizeMode = PictureBoxSizeMode.CenterImage; // задаем расположение изображения
        }
        Form2 form2;

        protected override void OnMouseWheel(MouseEventArgs e) // событие при прокрутки колеса мыши в picturebox
        {
            // Do not use MouseEventArgs.X, Y because they are relative! 
            Point pt_MouseAbs = Control.MousePosition;
            Control i_Ctrl = pictureBox1;
            do
            {
                Rectangle r_Ctrl = i_Ctrl.RectangleToScreen(i_Ctrl.ClientRectangle);
                if (!r_Ctrl.Contains(pt_MouseAbs))
                {
                    base.OnMouseWheel(e);
                    return; // mouse position is outside the picturebox or it's parents 
                }
                i_Ctrl = i_Ctrl.Parent;
            }
            while (i_Ctrl != null && i_Ctrl != this);

            //here code is
            if ((e.Delta < 0)&&((pictureBox1.Width<300)||(pictureBox1.Height<300)))// если изображение слишком маленькое
            {
                return; // выходим из функции
            }
            if ((e.Delta > 0)&& ((pictureBox1.Width > 1200) || (pictureBox1.Height > 3000))){//если изображение слишком большое
                return; // выходим из функции
            }
            
            pictureBox1.Width = pictureBox1.Width + (e.Delta / 10); // масштабируем изображение
            pictureBox1.Height = pictureBox1.Height + (e.Delta / 10);// масштабируем изображение

        }
        /*public Image ResizeImg(Image b, int nWidth, int nHeight)
        {
            pictureBox1.SizeMode = PictureBoxSizeMode.CenterImage;
            Image result = new Bitmap(nWidth, nHeight);
            using (Graphics g = Graphics.FromImage((Image)result))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(b, 0, 0, nWidth, nHeight);
                g.Dispose();
            }
            return result;
        }*/
        private void button1_Click(object sender, EventArgs e)
        {

            Scanner();//получение данных сканера из cfg-файла. Если данные неверны, включается режим ручной настройки сканера

            pictureBox1.Image = null;// удаляем всё из picturebox
            form2 = new Form2();
            this.Visible = false; //скрываем первую форму
            form2.Show();//показываем вторую форму
            var MS = new MemoryStream();
            MS = MemScan();// выполняем сканирование
            form2.Hide();// прячем вторую форму
            this.Visible = true;// делаем первую форму видимой
            this.Focus();// переводим фокус на первую форму

            if (MS != null)
            {
                try
                {
                    pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                    pictureBox1.Image = System.Drawing.Image.FromStream(MS);// загружаем изображение в picturebox
                }catch(Exception ex)
                {
                    MessageBox.Show("Изображение было отсканировано, но возникла ошибка вида: "+ex.Message);
                }
            }
            else
            {

                MessageBox.Show("Изображение не отсканировано. Возможно, сканер не подключен");

            }
        }



        public const string Config = "scanner.cfg";
        private Device _scanDevice;
        private Item _scannerItem;
        private Random _rnd = new Random();

        private Dictionary<string, object> _defaultDeviceProp;

        public bool IsVirtual;

        public void Scanner()
        {
            try
            {
                LoadConfig();// пытаемся загрузить конфигурацию из cfg-файла
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка конфигурации, требуется ручная настройка сканера");
                Configuration();// включаем режим ручной настройки сканера
            }
        }

        public void Configuration()//процедура ручной настройки сканера
        {
            try
            {
                var commonDialog = new CommonDialogClass();
                _scanDevice = commonDialog.ShowSelectDevice(WiaDeviceType.ScannerDeviceType, true);

                if (_scanDevice == null)
                    return;

                var items = commonDialog.ShowSelectItems(_scanDevice);

                if (items.Count < 1)
                    return;

                _scannerItem = items[1];

                SaveProp(_scanDevice.Properties, ref _defaultDeviceProp);

                SaveConfig();// сохраняем конфигурацию
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "Интерфейс сканера не доступен");
            }
        }

        private void SaveProp(WIA.Properties props, ref Dictionary<string, object> dic)
        {
            if (dic == null) dic = new Dictionary<string, object>();

            foreach (Property property in props)
            {
                var propId = property.PropertyID.ToString();
                var propValue = property.get_Value();

                dic[propId] = propValue;
            }
        }

        public void SetDuplexMode(bool isDuplex) // процедура установки режима двустороннего сканирования (не реализована)
        {
            // WIA property ID constants
            const string wiaDpsDocumentHandlingSelect = "3088";
            const string wiaDpsPages = "3096";

            // WIA_DPS_DOCUMENT_HANDLING_SELECT flags
            const int feeder = 0x001;
            const int duplex = 0x004;

            if (_scanDevice == null) return;

            if (isDuplex)
            {
                SetProp(_scanDevice.Properties, wiaDpsDocumentHandlingSelect, (duplex | feeder));
                SetProp(_scanDevice.Properties, wiaDpsPages, 1);
            }
            else
            {
                try
                {
                    SetProp(_scanDevice.Properties, wiaDpsDocumentHandlingSelect, _defaultDeviceProp[wiaDpsDocumentHandlingSelect]);
                    SetProp(_scanDevice.Properties, wiaDpsPages, _defaultDeviceProp[wiaDpsPages]);
                }
                catch (Exception e)
                {
                    MessageBox.Show(String.Format("Сбой восстановления режима сканирования:{0}{1}", Environment.NewLine, e.Message));
                }
            }
        }

        public MemoryStream MemScan()
        {
            //var s = new MemoryStream();

           
                if ((_scannerItem == null) && (!IsVirtual))
                {
                    MessageBox.Show("Сканер не настроен, обратитесь в меню 'Подготовка к сканированию' в разделе 'Информация'", "Info");
                    //return null;
                    //IsVirtual = true;
                }

                var stream = new MemoryStream();

                if (IsVirtual)
                {
                    if (_rnd.Next(3) == 0)
                    {
                        return null;
                    }

                    var btm = GetVirtualScan();
                    btm.Save(stream, ImageFormat.Jpeg);
                    return stream;
                }

                try
                {
                    var result = _scannerItem.Transfer(FormatID.wiaFormatJPEG);
                    var wiaImage = (ImageFile)result;
                    var imageBytes = (byte[])wiaImage.FileData.get_BinaryData();

                    using (var ms = new MemoryStream(imageBytes))
                    {
                        using (var bitmap = Bitmap.FromStream(ms))
                        {
                            bitmap.Save(stream, ImageFormat.Jpeg);
                        }
                    }

                }
                catch (Exception)
                {
                    return null;
                }

                return stream;

           
            
        }

        private Bitmap GetVirtualScan()// не нужно
        {
            const int imgSize = 777;
            var defBtm = new Bitmap(imgSize, imgSize);
            var g = Graphics.FromImage(defBtm);

            var r = new Random();

            g.FillRectangle(new SolidBrush(Color.FromArgb(r.Next(0, 50), r.Next(0, 50), r.Next(0, 50))), 0, 0, imgSize, imgSize); // bg

            for (int i = 0; i < r.Next(1000, 3000); i++)
            {
                var den = r.Next(200, 255);
                var br = new SolidBrush(Color.FromArgb(den, den, den));

                den -= 100;

                var pr = new Pen(Color.FromArgb(den, den, den), 1);

                var size = r.Next(1, 8);
                var x = r.Next(0, imgSize - size);
                var y = r.Next(0, imgSize - size);
                g.FillEllipse(br, x, y, size, size);
                g.DrawEllipse(pr, x, y, size, size);
            }

            g.DrawString("Виртуальный сканер", new Font(FontFamily.GenericSerif, 25), Brushes.Aqua, new RectangleF(0, 0, imgSize, imgSize));

            g.Flush();

            return defBtm;
        }

        private void SaveConfig()// сохранение конфигурации в файл
        {
            var settings = new List<string>();
            settings.Add("[device]");
            settings.Add(String.Format("DeviceID;{0}", _scanDevice.DeviceID));

            foreach (IProperty property in _scanDevice.Properties)
            {
                var propstring = string.Format("{1}{0}{2}{0}{3}", ";", property.Name, property.PropertyID, property.get_Value());
                settings.Add(propstring);
            }

            settings.Add("[item]");
            settings.Add(String.Format("ItemID;{0}", _scannerItem.ItemID));
            foreach (IProperty property in _scannerItem.Properties)
            {
                var propstring = string.Format("{1}{0}{2}{0}{3}", ";", property.Name, property.PropertyID, property.get_Value());
                settings.Add(propstring);
            }

            File.WriteAllLines(Config, settings.ToArray());
        }

        private enum loadMode { undef, device, item };

        private void LoadConfig()//загрузка конфигурации из файла
        {
            var settings = File.ReadAllLines(Config);

            var mode = loadMode.undef;

            foreach (var setting in settings)
            {
                if (setting.StartsWith("[device]"))
                {
                    mode = loadMode.device;
                    continue;
                }

                if (setting.StartsWith("[item]"))
                {
                    mode = loadMode.item;
                    continue;
                }

                if (setting.StartsWith("DeviceID"))
                {
                    var deviceid = setting.Split(';')[1];
                    var devMngr = new DeviceManagerClass();

                    foreach (IDeviceInfo deviceInfo in devMngr.DeviceInfos)
                    {
                        if (deviceInfo.DeviceID == deviceid)
                        {
                            _scanDevice = deviceInfo.Connect();
                            break;
                        }
                    }

                    if (_scanDevice == null)
                    {
                        MessageBox.Show("Сканнер из конигурации не найден");
                        return;
                    }

                    _scannerItem = _scanDevice.Items[1];
                    continue;
                }

                if (setting.StartsWith("ItemID"))
                {
                    var itemid = setting.Split(';')[1];
                    continue;
                }

                var sett = setting.Split(';');
                switch (mode)
                {
                    case loadMode.device:
                        SetProp(_scanDevice.Properties, sett[1], sett[2]);
                        break;

                    case loadMode.item:
                        SetProp(_scannerItem.Properties, sett[1], sett[2]);
                        break;
                }
            }
            SaveProp(_scanDevice.Properties, ref _defaultDeviceProp);
        }

        private static void SetProp(IProperties prop, object property, object value)
        {
            try
            {
                prop[property].set_Value(value);
            }
            catch (Exception)
            {
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)// сохранение картинки
        {
            if (pictureBox1.Image != null)
            {
                saveFileDialog1.FileName = "picture";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        switch (saveFileDialog1.FilterIndex)
                        {
                            case 1:
                                pictureBox1.Image.Save(saveFileDialog1.FileName, ImageFormat.Jpeg);
                                break;
                            case 2:
                                pictureBox1.Image.Save(saveFileDialog1.FileName, ImageFormat.Png);
                                break;
                            case 3:
                                pictureBox1.Image.Save(saveFileDialog1.FileName, ImageFormat.Bmp);
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(String.Format("Ошибка сохранения файла с аналогичным именем. Введите другое имя файла и попробуйте снова. Ошибка: " + ex.Message));
                        button2.PerformClick();
                    }
                    
                }
            }
        }

        private void сканироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button1.PerformClick();

        }

        private void настроитьСканерToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Configuration();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                pictureBox1.Image.RotateFlip(RotateFlipType.RotateNoneFlipX);
                pictureBox1.Refresh();
            }
        }
        // переворот и отражение изображения
        private void button5_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                pictureBox1.Image.RotateFlip(RotateFlipType.RotateNoneFlipX);
                pictureBox1.Refresh();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                int a;
                a = pictureBox1.Width;
                pictureBox1.Width = pictureBox1.Height;
                pictureBox1.Height = a;
                pictureBox1.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                pictureBox1.Refresh();
            }
        }

        /*private void button7_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                int a;
                a = pictureBox1.Width;
                pictureBox1.Width = pictureBox1.Height;
                pictureBox1.Height = a;
                pictureBox1.Image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                pictureBox1.Refresh();
            }
        }*/
        
        private void копироватьToolStripMenuItem_Click(object sender, EventArgs e)//копирование изображения в буфер обмена
        {
            try
            {
                if (pictureBox1.Image != null)
                {
                    Clipboard.SetImage(pictureBox1.Image);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Невозможно скопировать данное изображение! Ошибка: "+ex.Message);
            }
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)// при нажатии ПКМ
            {
                contextMenuStrip1.Show(this.Location);//показываем контекстное меню
            }
        }

        private void pictureBox1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void сохранитьКакToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button2.PerformClick();
        }

        private void печатьToolStripMenuItem_Click(object sender, EventArgs e)// выход из приложения
        {
            Application.Exit();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            try
            {
                pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    pictureBox1.Load(openFileDialog1.FileName);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(String.Format("Обнаружена ошибка несоответствия изображения. Ошибка: " + ex.Message));
            }
        }

        private void вставитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                if (Clipboard.GetImage() != null) pictureBox1.Image = Clipboard.GetImage();
            }
            catch(Exception ex)
            {
                MessageBox.Show(String.Format("Обнаружена ошибка несоответствия изображения. Ошибка: " + ex.Message));
            }
        }

        private void настроитьПринтерToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void printPreviewDialog1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)// печать изображения
        {
            try
            {
                if (pictureBox1.Image != null)
                {
                    System.Drawing.Printing.PrintDocument pd = new System.Drawing.Printing.PrintDocument();
                    pd.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printDocument1_PrintPage);
                    if (printDialog1.ShowDialog() == DialogResult.OK) pd.Print();
                }
                else
                {
                    MessageBox.Show("Перед печатью необходимо загрузить изображение!","Загрузите изображение!");
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка печати. Возможно, изображение не было загружено или имело неверный входной формат. Попробуйте заново загрузить либо отсканировать изображение.\n\n"+ex.Message,"Ошибка печати");
            }
        }
        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)// печать
        {
            try
            {
                e.Graphics.DrawImage(pictureBox1.Image, 10, 10); //(Standard paper size is 850 x 1100 or 2550 x 3300 pixels)
            }
            catch(Exception ex)
            {
                MessageBox.Show("Нет изображения для печати. Повторите загрузку изображения\n\n" + ex.Message, "Ошибка печати");
            }
        }

        private void button7_Click_1(object sender, EventArgs e)// поворот изображения
        {
            if (pictureBox1.Image != null)
            {
                int a;
                a = pictureBox1.Width;
                pictureBox1.Width = pictureBox1.Height;
                pictureBox1.Height = a;
                pictureBox1.Image.RotateFlip(RotateFlipType.Rotate90FlipNone);
                pictureBox1.Refresh();
            }
        }

        private void печатьToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            button3.PerformClick();
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("**** При первом запуске программы (ВАРИАНТ 1):****\n \n" +
                            "1. Подключите сканер к ПК и включите его." +
                            "2. В меню 'сканер' выберите пункт 'Настроить сканер'. Откроется диалоговое окно настройки.\n" +
                            "3. Выберите нужный сканер и нажмите кнопку 'ОК'\n" +
                            "4. Выберите тип сканируемых изображений и нажмите кнопку 'сканировать'\n \n" +
                            "ПРИМЕЧАНИЕ: сканирование НЕ начнется автоматически.\n" +
                            "В СЛУЧАЕ ОШИБКИ: проверьте, подключен ли сканер к компьютеру\n\n" +
                            "**** При первом запуске программы (ВАРИАНТ 2):****\n \n" +
                            "1. Нажмите кнопку 'сканирование'\n" +
                            "2. Запустится окно найтройки сканера\n" +
                            "3. Повторите пункты 1-4 из предыдущего раздела\n\n" +
                            "ПРИМЕЧАНИЕ: после настройки сканирование запустится автоматически\n" +
                            "В СЛУЧАЕ ОШИБКИ: проверьте, подключен ли сканер к компьютеру\n\n" +
                            "**** При повторном запуске программы:****\n \n" +
                            "Параметры сканера сохраняются автоматически после первой настройки." +
                            "При повторном использовании программы нажмите кнопку 'сканировать'\n \n" +
                            "В СЛУЧАЕ ОШИБКИ: произведите повторную настройку сканера (см. раздел 1) \n\n" +
                            "****Печать изображения****\n \n" +
                            "1. Нажмите кнопку 'печать'\n" +
                            "2. Выберите необходимый принтер\n" +
                            "3. Подтвердите печать", "Информация");
        }

        private void оПрограммеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Программа предназначена для сканирования и печати изображений, а так же, сохранения изображений в файл (из загрузки из файла)" +
                "в различных графических форматах.");
        }

        private void оРазработчикеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Олег Ярига (OlegYariga), 2018 год.");
        }
    }


}