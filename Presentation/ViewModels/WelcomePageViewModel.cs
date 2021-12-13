using System;

namespace Presentation.ViewModels
{
    /// <summary>      
    /// View Model of WelcomePage, responsible for logic for respected view.      
    /// </summary>
    public class WelcomePageViewModel
    {
        #region Properties      
        /// <summary>      
        /// This string property will have default text for demo purpose.    
        /// </summary>      
        private string _imGoodByeText = "This is binded from WelcomePageViewModel, Thank you for being part of this Blog!";
        /// <summary>      
        /// This string property will be binded with Textblock on view       
        /// </summary>      
        public string ImGoodByeText
        {
            get { return _imGoodByeText; }
            set { _imGoodByeText = value; }
        }
        #endregion
    }
}
