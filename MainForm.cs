using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;

namespace ChristiBase
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        # region const
        const string csACEConnStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\\Christina\\Students.accdb;";
        const string qSelMain = "Select MAIN.SN,DR,F,I,O,XSEX,DB,DOC,REG,ADDR from (((MAIN left outer join ADDRES on MAIN.SN=ADDRES.SN_IO) left outer join SEX on MAIN.SEX=SEX.ID) left outer join UKR on ADDRES.IDREG=UKR.ID)";
        const string qSelSchl = "Select REG,SCHL from ((MAIN left outer join SCHOOL on MAIN.SN=SCHOOL.SN_IO) left outer join UKR on SCHOOL.IDREG=UKR.ID) where SN_IO=";
        const string qSelSts = "Select XEDU,STS,RPRT from (((MAIN left outer join STSPERS on MAIN.SN=STSPERS.SN_IO) left outer join STATUS on STSPERS.IDSTS=STATUS.ID) left outer join EDUCAT on MAIN.EDU=EDUCAT.ID) where SN_IO=";
        const string qSelPay = "Select PAY.SN,MNH,PAY from ((MAIN left outer join PAY on MAIN.SN=PAY.SN_IO) left outer join MNTH on PAY.IDMNH=MNTH.ID) where SN_IO=";
        const string qSelSub = "Select SUBPERS.SN,SUB from ((MAIN left outer join SUBPERS on MAIN.SN=SUBPERS.SN_IO) left outer join SUBJECT on SUBPERS.IDSUB=SUBJECT.ID) where SN_IO=";
        const string strSqlSelMax = "Select max(SN) from Main";
        const string strSqlInsMain = "Insert into MAIN (DR,F,I,O,SEX,DB,DOC,EDU) values ('{0}','{1}','{2}','{3}',{4},'{5}','{6}',{7})";
        const string strSqlInsAddres = "Insert into ADDRES (SN,SN_IO,IDREG,ADDR) values ({0},{0},{1},'{2}')";
        const string strSqlInsSchool = "Insert into SCHOOL (SN,SN_IO,IDREG,SCHL) values ({0},{0},{1},'{2}')";
        const string strSqlInsStatus = "Insert into STSPERS (SN,SN_IO,IDSTS,RPRT) values ({0},{0},{1},'{2}')";
        const string strSqlInsPay = "Insert into PAY (SN_IO,IDMNH,PAY) values ({0},{1},'{2}')";
        const string strSqlInsSub = "Insert into SUBPERS (SN_IO,IDSUB) values ({0},{1})";
        const string strSqlUpdMain = "Update MAIN set DR='{1}',F='{2}',I='{3}',O='{4}',SEX={5},DB='{6}',DOC='{7}',EDU={8} where SN={0}";
        const string strSqlUpdAddres = "Update ADDRES set IDREG={1},ADDR='{2}' where SN_IO={0}";
        const string strSqlUpdSchool = "Update SCHOOL set IDREG={1},SCHL='{2}' where SN_IO={0}";
        const string strSqlUpdStatus = "Update STSPERS set IDSTS={1},RPRT='{2}' where SN_IO={0}";
        const string strSqlDel = "Delete from MAIN where SN=";
        const string strSqlDelPay = "Delete from PAY where SN=";
        const string strSqlDelSub = "Delete from SUBPERS where SN=";
        const string qMain = "MAIN";
        const string qSchl = "SCHL";
        const string qSts = "STS";
        const string qPay = "PAY";
        const string qSub = "SUB";

        # endregion

        # region var
        private string SNMain = string.Empty;
        private OleDbCommand cmd = new OleDbCommand();
        public OleDbConnection cn = new OleDbConnection();
        private DataSet DS = new DataSet();
        private OleDbDataAdapter Adptr = new OleDbDataAdapter();
        private OleDbDataAdapter Adptr2 = new OleDbDataAdapter();
        private BindingSource bsMain = new BindingSource();
        private BindingSource bsSchl = new BindingSource();
        private BindingSource bsSts = new BindingSource();
        private BindingSource bsPay = new BindingSource();
        private BindingSource bsSub = new BindingSource();


        # endregion





        //Подключение к БД
        private void MainForm_Load(object sender, EventArgs e)
        {
            dbConnect(cn);
            ShowList(qSelMain);
            dbDisConnect(cn);
        }
        
        public void dbConnect(OleDbConnection conn)
        {
            try
            {
                conn.ConnectionString = csACEConnStr;
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "Помилка підключення до БД",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void dbDisConnect(OleDbConnection conn)
        {
            try
            {
                conn.Close();
            }
            catch (OleDbException ex)
            {
                MessageBox.Show(ex.Message, "Помилка відключення від БД",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        void ShowList(string qSQL)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                //Adptr.SelectCommand  = new OleDbCommand(string.Format(qSelAll, qMain), cn);
                //Adptr.SelectCommand = new OleDbCommand(qSQL, cn);
                //if (DS.Tables.Contains(qMain) == false) DS.Tables.Add(qMain);
                //DS.Tables[qMain].Clear();
                //bsMain.DataSource = DS.Tables[qMain];
                //Adptr.Fill(DS, qMain);
                cmd.Connection = cn;
                cmd.Connection.Open(); 
                cmd.CommandText = qSQL;
                bsMain.Clear();
                bsMain.DataSource = cmd.ExecuteReader();
                cmd.Connection.Close();
                if (bsMain.Count > 0) 
                {
                    dgvMain.DataSource = bsMain;
                    dgvMain.Columns[0].Visible = false;
                    dgvMain.Columns[1].HeaderText = "Дата реєстрації";
                    dgvMain.Columns[2].HeaderText = "Прізвище";
                    dgvMain.Columns[3].HeaderText = "Ім'я";
                    dgvMain.Columns[4].HeaderText = "По-батькові";
                    dgvMain.Columns[5].HeaderText = "Стать";
                    dgvMain.Columns[6].HeaderText = "Дата народження";
                    dgvMain.Columns[7].HeaderText = "Паспорт";
                    dgvMain.Columns[8].HeaderText = "Регіон";
                    dgvMain.Columns[9].HeaderText = "Місце проживання";
                    dgvMain.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[5].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[6].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[7].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[8].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    dgvMain.Columns[9].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                }
                dgvMain.Select();
                statusStrip.Items[0].Text = "Всього відібрано "+dgvMain.RowCount.ToString()+" слухачів(ча) підготовчого відділення";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка відбору слухачів",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        void RefreshSchl()
        {
            cmd.Connection = cn;
            cmd.Connection.Open();
            cmd.CommandText = qSelSchl + SNMain;
            bsSchl.DataSource = cmd.ExecuteReader();
            cmd.Connection.Close();             
            dgvSchl.DataSource = bsSchl;
            dgvSchl.Columns[0].HeaderText = "Регіон";
            dgvSchl.Columns[1].HeaderText = "Навчальний заклад";
            dgvSchl.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvSchl.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        void RefreshSts()
        {
            cmd.Connection = cn;
            cmd.Connection.Open(); 
            cmd.CommandText = qSelSts + SNMain;
            bsSts.DataSource = cmd.ExecuteReader();
            cmd.Connection.Close(); 
            dgvSts.DataSource = bsSts;
            dgvSts.Columns[0].HeaderText = "Форма навчання";
            dgvSts.Columns[1].HeaderText = "Статус";
            dgvSts.Columns[2].HeaderText = "Наказ";
            dgvSts.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvSts.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            dgvSts.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        }

        void RefreshPay()
        {
            cmd.Connection = cn;
            cmd.Connection.Open(); 
            cmd.CommandText = qSelPay + SNMain;
            bsPay.DataSource = cmd.ExecuteReader();
            cmd.Connection.Close(); 
            if (bsPay.Count > 0)
            {
                dgvPay.DataSource = bsPay; 
                dgvPay.Columns[0].Visible = false;
                dgvPay.Columns[1].HeaderText = "Місяць";
                dgvPay.Columns[2].HeaderText = "Оплата";
                dgvPay.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                dgvPay.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            }
        }

        void RefreshSub()
        {
            cmd.Connection = cn;
            cmd.Connection.Open();
            cmd.CommandText = qSelSub + SNMain;
            bsSub.DataSource = cmd.ExecuteReader();
            cmd.Connection.Close(); 
            if (bsSub.Count > 0)
            {
                dgvSub.DataSource = bsSub; 
                dgvSub.Columns[0].Visible = false;
                dgvSub.Columns[1].HeaderText = "Предмети";
                dgvSub.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

            }
        }

        private void dgvMain_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dgvMain.RowCount > 0) 
            {
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    cmd.Connection.Close(); 
                    SNMain = dgvMain.Rows[e.RowIndex].Cells["SN"].Value.ToString();
                    RefreshSchl();
                    RefreshSts();
                    RefreshPay();
                    RefreshSub();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Помилка відбору реквізитів слухачів",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        void addDgvMain(string newDR, string newF, string newI, string newO, string newSex, 
                        string newDB, string newDoc, string newRegAddr, string newAddr, 
                        string newRegSchl, string newSchl, string newEdu, string newSts, string newRprt)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;       
                cmd.Connection = cn;
                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlInsMain, newDR, newF, newI, newO, newSex, newDB, newDoc, newEdu);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = strSqlSelMax;
                string maxSN = cmd.ExecuteScalar().ToString();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlInsAddres, maxSN, newRegAddr, newAddr);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlInsSchool, maxSN, newRegSchl, newSchl);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlInsStatus, maxSN, newSts, newRprt);
                cmd.ExecuteReader();
                cmd.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка додавання слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        void updDgvMain(string newDR, string newF, string newI, string newO, string newSex, 
                        string newDB, string newDoc, string newRegAddr, string newAddr, 
                        string newRegSchl, string newSchl, string newEdu, string newSts, string newRprt)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor; 
                cmd.Connection = cn;
                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlUpdMain, SNMain, newDR, newF, newI, newO, newSex, newDB, newDoc, newEdu);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlUpdAddres, SNMain, newRegAddr, newAddr);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlUpdSchool, SNMain, newRegSchl, newSchl);
                cmd.ExecuteReader();
                cmd.Connection.Close();

                cmd.Connection.Open();
                cmd.CommandText = string.Format(strSqlUpdStatus, SNMain, newSts, newRprt);
                cmd.ExecuteReader();
                cmd.Connection.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Помилка зміни реквізитів слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        string BuildSql(string Col, string Val)
        {
            string s;
            if (string.IsNullOrEmpty(Val.Trim())) s = string.Empty;
            else s = Col + "=" + Val;
            return s;
        }

        string Qut(string s)
        {
            return s = "'" + s + "'";
        }

        void findDgvMain(string newDR, string newF, string newI, string newO, string newSex,
                         string newDB, string newDoc, string newRegAddr, string newAddr,
                         string newRegSchl, string newSchl, string newEdu, string newSts, string newRprt)
        {
            this.Cursor = Cursors.WaitCursor; 
            string sql = string.Empty;
            sql = BuildSql("F", Qut(newF));
            if (string.IsNullOrEmpty(sql)) sql = BuildSql("I", Qut(newI));
            else sql = sql + " and " + BuildSql("I", Qut(newI));
            //if (string.IsNullOrEmpty(sql)) sql = BuildSql("O", newO);
            //                          else sql = sql + " and " + BuildSql("O", newO);
            //if (string.IsNullOrEmpty(sql)) sql = BuildSql("Sex", newSex);
            //                          else sql = sql + " and " + BuildSql("Sex", newSex);
            //if (string.IsNullOrEmpty(sql)) sql = BuildSql("DB", newDB);
            //                          else sql = sql + " and " + BuildSql("DB", newDB);
            //if (string.IsNullOrEmpty(sql)) sql = BuildSql("EDU", newEdu);
            //                          else sql = sql + " and " + BuildSql("DOC", newDoc);
            //if (string.IsNullOrEmpty(sql)) sql = BuildSql("IDREG", newRegAddr);
            //                          else sql = sql + " and " + BuildSql("IDREG", newRegAddr);
            //if (string.IsNullOrEmpty(sql)) sql = BuildSql("ADDR", newAddr);
            //                          else sql = sql + " and " + BuildSql("ADDR", newAddr);

            if (!string.IsNullOrEmpty(sql)) sql = " where " + sql;
            this.Cursor = Cursors.Default; 
            ShowList(qSelMain + sql);
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            AddForm fmAdd = new AddForm();
            if (fmAdd.ShowDialog(this) == DialogResult.OK)
            {
                addDgvMain(fmAdd.newDR, fmAdd.newF, fmAdd.newI, fmAdd.newO, fmAdd.newSex,
                           fmAdd.newDB, fmAdd.newDoc, fmAdd.newRegAddr, fmAdd.newAddr,
                           fmAdd.newRegSchl, fmAdd.newSchl, fmAdd.newEdu, fmAdd.newSts, fmAdd.newRprt);
                ShowList(qSelMain);
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0) 
            {
                AddForm fmAdd = new AddForm();
                fmAdd.dtpDR.Text = dgvMain.CurrentRow.Cells[1].Value.ToString();
                fmAdd.txbF.Text = dgvMain.CurrentRow.Cells[2].Value.ToString();
                fmAdd.txbI.Text = dgvMain.CurrentRow.Cells[3].Value.ToString();
                fmAdd.txbO.Text = dgvMain.CurrentRow.Cells[4].Value.ToString();
                fmAdd.curSex = dgvMain.CurrentRow.Cells[5].Value.ToString();
                fmAdd.dtpDB.Text = dgvMain.CurrentRow.Cells[6].Value.ToString();
                fmAdd.txbDoc.Text = dgvMain.CurrentRow.Cells[7].Value.ToString();
                fmAdd.curAddr = dgvMain.CurrentRow.Cells[8].Value.ToString();
                fmAdd.txbAddr.Text = dgvMain.CurrentRow.Cells[9].Value.ToString();
                fmAdd.curSchl = dgvSchl.CurrentRow.Cells[0].Value.ToString();
                fmAdd.txbSchl.Text = dgvSchl.CurrentRow.Cells[1].Value.ToString();
                fmAdd.curEdu = dgvSts.CurrentRow.Cells[0].Value.ToString();
                fmAdd.curSts = dgvSts.CurrentRow.Cells[1].Value.ToString();
                fmAdd.txbRprt.Text = dgvSts.CurrentRow.Cells[2].Value.ToString();
                if (fmAdd.ShowDialog(this) == DialogResult.OK)
                {
                    updDgvMain(fmAdd.newDR, fmAdd.newF, fmAdd.newI, fmAdd.newO, fmAdd.newSex,
                               fmAdd.newDB, fmAdd.newDoc, fmAdd.newRegAddr, fmAdd.newAddr,
                               fmAdd.newRegSchl, fmAdd.newSchl, fmAdd.newEdu, fmAdd.newSts, fmAdd.newRprt);
                    ShowList(qSelMain);
                }
            } else MessageBox.Show("Немає жодного слухача для редагування!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                if (MessageBox.Show("Ви дійсно бажаєте видалити вибраного слухача?", "Попередження",
                                     MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = strSqlDel + SNMain;
                        cmd.ExecuteReader();
                        cmd.Connection.Close(); 
                        ShowList(qSelMain);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка видалення слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для видалення!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            AddForm fmAdd = new AddForm();
            fmAdd.txbF.Enabled = false;     fmAdd.chbF.Visible = true;
            fmAdd.txbI.Enabled = false;     fmAdd.chbI.Visible = true;
            fmAdd.txbO.Enabled = false;     fmAdd.chbO.Visible = true;
            fmAdd.dtpDB.Enabled = false;    fmAdd.chbDB.Visible = true;
            fmAdd.dtpDR.Enabled = false;    fmAdd.chbDR.Visible = true; 
            fmAdd.cmbSex.Enabled = false;   fmAdd.chbSex.Visible = true;
            fmAdd.txbDoc.Enabled = false;   fmAdd.chbDoc.Visible = true;
            fmAdd.cmbAddr.Enabled = false;
            fmAdd.txbAddr.Enabled = false;  fmAdd.chbAddr.Visible = true;
            fmAdd.cmbSchl.Enabled = false;
            fmAdd.txbSchl.Enabled = false;  fmAdd.chbSchl.Visible = true;
            fmAdd.cmbEdu.Enabled = false;   fmAdd.chbEdu.Visible = true;
            fmAdd.cmbSts.Enabled = false;
            fmAdd.txbRprt.Enabled = false;  fmAdd.chbSts.Visible = true;
            fmAdd.Find = true;
            MessageBox.Show("Пошук тимчасово не працює (окрім Прізвища + Ім'я) !!!", "Увага!",
                             MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if (fmAdd.ShowDialog(this) == DialogResult.OK)
            {
                findDgvMain(fmAdd.newDR, fmAdd.newF, fmAdd.newI, fmAdd.newO, fmAdd.newSex,
                            fmAdd.newDB, fmAdd.newDoc, fmAdd.newRegAddr, fmAdd.newAddr,
                            fmAdd.newRegSchl, fmAdd.newSchl, fmAdd.newEdu, fmAdd.newSts, fmAdd.newRprt);
            }
        }
        
        private void btnAddSub_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                AddSubForm fmAddSub = new AddSubForm();
                if (fmAddSub.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = string.Format(strSqlInsSub, SNMain, fmAddSub.newSub);
                        cmd.ExecuteReader();
                        cmd.Connection.Close(); 
                        RefreshSub();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка додавання предмета слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для додавання предмета!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnAddPay_Click(object sender, EventArgs e)
        {
            if (dgvMain.RowCount > 0)
            {
                AddPayForm fmAddPay = new AddPayForm();
                if (fmAddPay.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = string.Format(strSqlInsPay, SNMain, fmAddPay.newMnthPay, fmAddPay.newPay);
                        cmd.ExecuteReader();
                        cmd.Connection.Close();
                        RefreshPay();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка додавання оплати слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для додавання оплати!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnDelSub_Click(object sender, EventArgs e)
        {
            if (dgvSub.RowCount > 0)
            {
                if (MessageBox.Show("Ви дійсно бажаєте видалити вибраний предмет?", "Попередження",
                                     MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = strSqlDelSub + dgvSub.CurrentRow.Cells[0].Value.ToString();
                        cmd.ExecuteReader();
                        cmd.Connection.Close(); 
                        RefreshSub();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка видалення предмета слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для видалення предмета!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
   
        private void btnDelPay_Click(object sender, EventArgs e)
        {
            if (dgvPay.RowCount > 0)
            {
                if (MessageBox.Show("Ви дійсно бажаєте видалити вибрану оплату?", "Попередження",
                                     MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
                {
                    try
                    {
                        this.Cursor = Cursors.WaitCursor;
                        cmd.Connection = cn;
                        cmd.Connection.Open();
                        cmd.CommandText = strSqlDelPay + dgvPay.CurrentRow.Cells[0].Value.ToString();
                        cmd.ExecuteReader();
                        cmd.Connection.Close(); 
                        RefreshPay();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Помилка видалення оплати слухача", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        this.Cursor = Cursors.Default;
                    }
                }
            } else MessageBox.Show("Немає жодного слухача для видалення оплати!", "Попередження",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (MessageBox.Show("Ви дійсно бажаєте вийти з інформаційної системи?", "Попередження",
                                 MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.Cancel)
            {e.Cancel = true;}
        }

        //private void releaseObject(object obj)
        //{
        //    try
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        //        obj = null;
        //    }
        //    catch (Exception ex)
        //    {
        //        obj = null;
        //        MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
        //    }
        //    finally
        //    {
        //        GC.Collect();
        //    }
        //}
        //private void button1_Click(object sender, EventArgs e)
        //{
        //    Excel.Application xlApp;
        //    Excel.Workbook xlWorkBook;
        //    Excel.Worksheet xlWorkSheet;
        //    object misValue = System.Reflection.Missing.Value;

        //    xlApp = new Excel.ApplicationClass();
        //    xlWorkBook = xlApp.Workbooks.Add(misValue);
        //    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
        //    int i = 0;
        //    int j = 0;

        //    for (i = 0; i <= dgvMain.RowCount - 1; i++)
        //    {
        //        for (j = 0; j <= dgvMain.ColumnCount - 1; j++)
        //        {
        //            DataGridViewCell cell = dgvMain[j, i];
        //            xlWorkSheet.Cells[i + 1, j + 1] = cell.Value;
        //        }
        //    }

        //    xlWorkBook.SaveAs("csharp.net-informations.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        //    xlWorkBook.Close(true, misValue, misValue);
        //    xlApp.Quit();

        //    releaseObject(xlWorkSheet);
        //    releaseObject(xlWorkBook);
        //    releaseObject(xlApp);

        //    MessageBox.Show("Excel file created , you can find the file c:\\csharp.net-informations.xls");

        //}
      }

}
