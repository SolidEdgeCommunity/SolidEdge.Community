using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeCommunity.Extensions
{
    /// <summary>
    /// SolidEdgeFramework.Environment extension methods.
    /// </summary>
    public static class EnvironmentExtensions
    {
        /// <summary>
        /// Returns a Guid representing the environment category.
        /// </summary>
        public static Guid GetCategoryId(this SolidEdgeFramework.Environment environment)
        {
            return new Guid(environment.CATID);
        }

        /// <summary>
        /// Returns a Type representing the corresponding command constants for a particular environment.
        /// </summary>
        /// <param name="environment"></param>
        /// <returns></returns>
        public static Type GetCommandConstantType(this SolidEdgeFramework.Environment environment)
        {
            var categoryId = environment.GetCategoryId();

            if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Application))
            {
                return typeof(SolidEdgeConstants.SolidEdgeCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Assembly))
            {
                return typeof(SolidEdgeConstants.AssemblyCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.DMAssembly))
            {
                return typeof(SolidEdgeConstants.AssemblyCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.CuttingPlaneLine))
            {
                return typeof(SolidEdgeConstants.CuttingPlaneLineCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Draft))
            {
                return typeof(SolidEdgeConstants.DetailCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.DrawingViewEdit))
            {
                return typeof(SolidEdgeConstants.DrawingViewEditCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Explode))
            {
                return typeof(SolidEdgeConstants.ExplodeCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Layout))
            {
                return typeof(SolidEdgeConstants.LayoutCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Sketch))
            {
                return typeof(SolidEdgeConstants.LayoutInPartCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Motion))
            {
                return typeof(SolidEdgeConstants.MotionCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Part))
            {
                return typeof(SolidEdgeConstants.PartCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.DMPart))
            {
                return typeof(SolidEdgeConstants.PartCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Profile))
            {
                return typeof(SolidEdgeConstants.ProfileCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.ProfileHole))
            {
                return typeof(SolidEdgeConstants.ProfileHoleCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.ProfilePattern))
            {
                return typeof(SolidEdgeConstants.ProfilePatternCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.ProfileRevolved))
            {
                return typeof(SolidEdgeConstants.ProfileRevolvedCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.SheetMetal))
            {
                return typeof(SolidEdgeConstants.SheetMetalCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.DMSheetMetal))
            {
                return typeof(SolidEdgeConstants.SheetMetalCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Simplify))
            {
                return typeof(SolidEdgeConstants.SimplifyCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Studio))
            {
                return typeof(SolidEdgeConstants.StudioCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.XpresRoute))
            {
                return typeof(SolidEdgeConstants.TubingCommandConstants);
            }
            else if (categoryId.Equals(SolidEdgeSDK.EnvironmentCategories.Weldment))
            {
                return typeof(SolidEdgeConstants.WeldmentCommandConstants);
            }

            return null;
        }

        /// <summary>
        /// Returns a SolidEdgeFramework.Environment by name.
        /// </summary>
        public static SolidEdgeFramework.Environment LookupByName(this SolidEdgeFramework.Environments environments, string name)
        {
            for (int i = 1; i <= environments.Count; i++)
            {
                var environment = environments.Item(i);
                if (environment.Name.Equals(name))
                {
                    return environment;
                }
            }

            return null;
        }

        /// <summary>
        /// Returns a SolidEdgeFramework.Environment by category id..
        /// </summary>
        public static SolidEdgeFramework.Environment LookupByCategoryId(this SolidEdgeFramework.Environments environments, Guid categoryId)
        {
            for (int i = 1; i <= environments.Count; i++)
            {
                var environment = environments.Item(i);
                if (environment.GetCategoryId().Equals(categoryId))
                {
                    return environment;
                }
            }

            return null;
        }
    }
}
