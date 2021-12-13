using Prism.Ioc;
using Prism.Modularity;
using Prism.Regions;

namespace Presentation
{
    /// <summary>        
    /// Responsible for mapping modules        
    /// </summary> 
    public class ModuleLocators : IModule
    {
        #region private properties        
        /// <summary>        
        /// Instance of IRegionManager        
        /// </summary>    
        private IRegionManager _regionManager;
        public ModuleLocators(IRegionManager regionManager)
        {
            _regionManager = regionManager;
        }

        #endregion

        #region Interface methods        
        /// <summary>        
        /// Initializes Welcome page of your application.        
        /// </summary>        
        /// <param name="containerProvider"></param>  
        public void OnInitialized(IContainerProvider containerProvider)
        {
            _regionManager.RegisterViewWithRegion("Shell", typeof(View.WelcomePageView));  //ModuleLocators is added for testing purpose,     
                                                                                           //later we'll replace it with WelcomePageView      
        }

        /// <summary>        
        /// RegisterTypes used to register modules        
        /// </summary>        
        /// <param name="containerRegistry"></param>        
        public void RegisterTypes(IContainerRegistry containerRegistry)
        {
        }
        #endregion
    }
}
