using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SolidEdgeFramework.Extensions
{
    public static class EnvironmentExtensions
    {
        public static Guid GetCategoryId(this SolidEdgeFramework.Environment environment)
        {
            return new Guid(environment.CATID);
        }

        public static Type GetCommandConstantType(this SolidEdgeFramework.Environment environment)
        {
            var categoryId = environment.GetCategoryId();

            if (categoryId.Equals(SolidEdge.CATID.SEApplicationGuid))
            {
                return typeof(SolidEdgeConstants.SolidEdgeCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEAssemblyGuid))
            {
                return typeof(SolidEdgeConstants.AssemblyCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEDMAssemblyGuid))
            {
                return typeof(SolidEdgeConstants.AssemblyCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SECuttingPlaneLineGuid))
            {
                return typeof(SolidEdgeConstants.CuttingPlaneLineCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEDraftGuid))
            {
                return typeof(SolidEdgeConstants.DetailCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEDrawingViewEditGuid))
            {
                return typeof(SolidEdgeConstants.DrawingViewEditCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEExplodeGuid))
            {
                return typeof(SolidEdgeConstants.ExplodeCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SELayoutGuid))
            {
                return typeof(SolidEdgeConstants.LayoutCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SESketchGuid))
            {
                return typeof(SolidEdgeConstants.LayoutInPartCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEMotionGuid))
            {
                return typeof(SolidEdgeConstants.MotionCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEPartGuid))
            {
                return typeof(SolidEdgeConstants.PartCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEDMPartGuid))
            {
                return typeof(SolidEdgeConstants.PartCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEProfileGuid))
            {
                return typeof(SolidEdgeConstants.ProfileCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEProfileHoleGuid))
            {
                return typeof(SolidEdgeConstants.ProfileHoleCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEProfilePatternGuid))
            {
                return typeof(SolidEdgeConstants.ProfilePatternCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEProfileRevolvedGuid))
            {
                return typeof(SolidEdgeConstants.ProfileRevolvedCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SESheetMetalGuid))
            {
                return typeof(SolidEdgeConstants.SheetMetalCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEDMSheetMetalGuid))
            {
                return typeof(SolidEdgeConstants.SheetMetalCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SESimplifyGuid))
            {
                return typeof(SolidEdgeConstants.SimplifyCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEStudioGuid))
            {
                return typeof(SolidEdgeConstants.StudioCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEXpresRouteGuid))
            {
                return typeof(SolidEdgeConstants.TubingCommandConstants);
            }
            else if (categoryId.Equals(SolidEdge.CATID.SEWeldmentGuid))
            {
                return typeof(SolidEdgeConstants.WeldmentCommandConstants);
            }

            return null;
        }

        public static SolidEdgeFramework.Environment LookupByCategoryId(this SolidEdgeFramework.Environments environments, string name)
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

        public static SolidEdgeFramework.Environment LookupByName(this SolidEdgeFramework.Environments environments, Guid categoryId)
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
