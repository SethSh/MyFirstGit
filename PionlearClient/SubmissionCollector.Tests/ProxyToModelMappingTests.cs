using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SubmissionCollector.Tests
{
    [TestClass]
    public class ProxyToModelMappingTests
    {
        [TestMethod]
        public void PolicyRiskProfileProxyMapsToPolicyRiskProfile()
        {
            //var policyRiskProfileProxy = new PolicyRiskProfileProxy("ignore");

            //var rangeContent = new object[, ]
            //{
            //    {"Limit", "Attachment", "Percent"},
            //    {1e6, 0, 0.25},
            //    {1e7, 0, 0.75},
            //};

            //const RiskProfileWeightDescription weightDescription = RiskProfileWeightDescription.Percent;

            //List<string> dataTypeErrors;
            //var actualRiskProfile = policyRiskProfileProxy.RiskProfileMapper.Map(policyRiskProfileProxy, rangeContent, weightDescription, out dataTypeErrors);

            //var expectedRiskProfile = new RiskProfileModel
            //{
            //    Name = RiskProfileCategory.Policy,
            //    WeightDescription = RiskProfileWeightDescription.Percent,
            //    Rows = new List<RiskProfileRow>
            //    {
            //        new RiskProfileRow
            //        {
            //            Dimensions = new List<RowDimension> { new LimitDimension {Value = 1e6}, new SirDimension {Value = 0} },
            //            Data = 0.25
            //        },
            //        new RiskProfileRow
            //        {
            //            Dimensions = new List<RowDimension> { new LimitDimension {Value = 1e7}, new SirDimension {Value = 0} },
            //            Data = 0.75
            //        }
            //    }
            //};

            //MathAssertions.AreEqual(actualRiskProfile, expectedRiskProfile);
        }
        
        [TestMethod]
        public void HazardRiskProfileProxyMapsToHazardRiskProfile()
        {
            //var hazardRiskProfileProxy = new HazardRiskProfileProxy("ignore");

            //var rangeContent = new object[,]
            //{
            //    {"Hazard Name", "Percent"},
            //    {"Light & Medium", 0.25},
            //    {"Heavy", 0.75},
            //};

            //var mapper =  AutoLiabilityCommercialStringToCodeMapper.HazardDictionary;
            //hazardRiskProfileProxy.ReplaceStringsWithIntegerCodes(mapper, rangeContent);

            //const RiskProfileWeightDescription weightDescription = RiskProfileWeightDescription.Percent;

            //List<string> dataTypeErrors;
            //var actualRiskProfile = hazardRiskProfileProxy.RiskProfileMapper.Map(hazardRiskProfileProxy, rangeContent,
            //    weightDescription, out dataTypeErrors);

            //var expectedRiskProfile = new RiskProfileModel
            //{
            //    Name = RiskProfileCategory.Hazard,
            //    WeightDescription = RiskProfileWeightDescription.Percent,
            //    Rows = new List<RiskProfileRow>
            //    {
            //        new RiskProfileRow
            //        {
            //            Dimensions = new List<RowDimension> { new HazardDimension {Value = 18} },
            //            Data = 0.25
            //        },
            //        new RiskProfileRow
            //        {
            //            Dimensions = new List<RowDimension> { new HazardDimension { Value = 41} },
            //            Data = 0.75
            //        }
            //    }
            //};

            //MathAssertions.AreEqual(actualRiskProfile, expectedRiskProfile);
        }

        [TestMethod]
        public void StateRiskProfileProxyMapsToStateRiskProfile()
        {
            //var stateRiskProfileProxy = new StateRiskProfileProxy("ignore");

            //var rangeContent = new object[,]
            //{
            //    {"State Abbrev", "State Name", "State Group", "St Group Public", "Percent"},
            //    {"NJ", "New Jersey", "Group1", "Group1", 0.25},
            //    {"AL", "Alabama", "GroupX", "GroupX", 0.70},
            //    {"CW", "Countrywide", "GroupX", "GroupX", 0.05},
            //};

            //StateCodesFromBex.GetStateReferenceData();
            //var mapper = StateCodesFromBex.ReferenceData;

            //stateRiskProfileProxy.ReplaceStringsWithIntegerCodes(mapper, rangeContent);
            //const RiskProfileWeightDescription weightDescription = RiskProfileWeightDescription.Percent;

            //List<string> dataTypeErrors;
            //var actualRiskProfile = stateRiskProfileProxy.RiskProfileMapper.Map(stateRiskProfileProxy, rangeContent, weightDescription, out dataTypeErrors);

            //var expectedRiskProfile = new RiskProfileModel
            //{
            //    Name = RiskProfileCategory.State,
            //    WeightDescription = RiskProfileWeightDescription.Percent,
            //    Rows = new List<RiskProfileRow>
            //    {
            //        new RiskProfileRow
            //        {
            //            Dimensions = new List<RowDimension> { new StateDimension {Value = 32} },
            //            Data = 0.25
            //        },
            //        new RiskProfileRow
            //        {
            //            Dimensions = new List<RowDimension> { new StateDimension { Value = 2} },
            //            Data = 0.70
            //        },
            //        new RiskProfileRow
            //        {
            //            Dimensions = new List<RowDimension> { new StateDimension { Value = null} },
            //            Data = 0.05
            //        }
            //    }
            //};

            //MathAssertions.AreEqual(actualRiskProfile, expectedRiskProfile);
        }
    }
}

                  

