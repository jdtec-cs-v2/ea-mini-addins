using EA;
using System;

namespace MiniAddinsFacade
{
    /// <summary>
    /// EA Add-ins Entry Module
    /// </summary>
    public class Main
	{
		private bool m_ShowFullMenus = false;
        private frmModel frmFacade;
        private EA.ModelWatcher modelWatcher;
        
        public String EA_Connect(EA.Repository Repository) 
		{
			// No special processing req'd
			return "";
		}

        /// <summary>
        /// EA侧注册插件菜单项
        /// </summary>
        /// <param name="Repository"></param>
        /// <param name="Location"></param>
        /// <param name="MenuName"></param>
        /// <returns></returns>
		public object EA_GetMenuItems(EA.Repository Repository, string Location, string MenuName) 
		{
			
			switch( MenuName )
			{
				case "":
					return "-&Mini Add-ins";
				case "-&Mini Add-ins":
					string[] ar = { "&Export Image", "&Export Logical Model", "&Statistics Report", "-", "&Export Image(with modeless)", "-", "Show All &Options" };
					return ar;
			}
			return "";
		}

        /// <summary>
        /// 判断是否打开了EA工程
        /// </summary>
        /// <param name="Repository"></param>
        /// <returns></returns>
		bool IsProjectOpen(EA.Repository Repository)
		{
			try
			{
				EA.Collection c = Repository.Models;
				return true;
			}
			catch
			{
				return false;
			}
		}

        /// <summary>
        /// 设置菜单状态，EA侧查询用
        /// </summary>
        /// <param name="Repository"></param>
        /// <param name="Location"></param>
        /// <param name="MenuName"></param>
        /// <param name="ItemName"></param>
        /// <param name="IsEnabled"></param>
        /// <param name="IsChecked"></param>
		public void EA_GetMenuState(EA.Repository Repository, string Location, string MenuName, string ItemName, ref bool IsEnabled, ref bool IsChecked)
		{
			if( IsProjectOpen(Repository) )
			{
				if( ItemName == "Show All &Options" )
					IsChecked = m_ShowFullMenus;
				else if( ItemName == "&Statistics Report")
					IsEnabled = m_ShowFullMenus;
			}
			else
            {
                // If no open project, disable all menu options
                IsEnabled = false;
            }
				
		}

        /// <summary>
        /// Add-ins 菜单调用
        /// </summary>
        /// <param name="Repository"></param>
        /// <param name="Location"></param>
        /// <param name="MenuName"></param>
        /// <param name="ItemName"></param>
		public void EA_MenuClick(EA.Repository Repository, string Location, string MenuName, string ItemName)
		{
			
			switch( ItemName )
			{
				case "Show All &Options":

                    m_ShowFullMenus = !m_ShowFullMenus;
					break;

				case "&Export Image":

                    frmModel frmFacadeModel = new frmModel();
                    frmFacadeModel.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                    frmFacadeModel.setEARepository(Repository);
                    frmFacadeModel.SetDisplayToolCatsIndex(0);
                    frmFacadeModel.OpenForm(0);
                    break;

                case "&Export Logical Model":

                    frmModel frmFacadeLDMModel = new frmModel();
                    frmFacadeLDMModel.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                    frmFacadeLDMModel.setEARepository(Repository);
                    frmFacadeLDMModel.SetDisplayToolCatsIndex(2);
                    frmFacadeLDMModel.OpenForm(0);
                    break;

                case "&Export Image(with modeless)":

                    if(this.frmFacade == null)
                    {
                        frmFacade = new frmModel();
                        frmFacade.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                        this.modelWatcher = Repository.CreateModelWatcher();
                    }
                    
                    frmFacade.setEARepository(Repository);
                    frmFacade.setEAModelWatcher(this.modelWatcher);
                    frmFacade.SetDisplayToolCatsIndex(0);
                    frmFacade.OpenForm(1);
                    break;

                case "&Statistics Report":

                    frmModel frmFacadeModelOther = new frmModel();
                    frmFacadeModelOther.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
                    frmFacadeModelOther.setEARepository(Repository);
                    frmFacadeModelOther.SetDisplayToolCatsIndex(1);
                    frmFacadeModelOther.OpenForm(0);
                    break;
			}
		}

        /// <summary>
        /// Add-ins关闭时
        /// </summary>
        public void EA_Disconnect()  
		{
			GC.Collect();  
			GC.WaitForPendingFinalizers();   
		}
	}
}
