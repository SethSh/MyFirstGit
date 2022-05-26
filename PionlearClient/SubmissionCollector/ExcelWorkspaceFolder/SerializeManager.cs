using System;
using System.Diagnostics;
using System.Text;
using Newtonsoft.Json;
using PionlearClient;
using PionlearClient.Extensions;
using PionlearClient.TokenAuthentication;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal static class SerializeManager
    {
        public static string ConvertToJson<T>(T obj)
        {
            return JsonConvert.SerializeObject(
                obj,
                Formatting.Indented,
                new JsonSerializerSettings
                {
                    TypeNameHandling = TypeNameHandling.Auto,
                    PreserveReferencesHandling = PreserveReferencesHandling.Objects
                });
        }

        public static void ConvertModelToJson(IPackage package, IWorkbookLogger logger)
        {
            try
            {
                using (new CursorToWait())
                {
                    var packageModel = package.CreatePackageModel();

                    if (package.ExcelValidation.Length > 0)
                    {
                        MessageHelper.Show(BexConstants.DataValidationTitle, package.ExcelValidation.ToString(), MessageType.Stop);
                    }

                    var validation = packageModel.Validate();
                    if (validation.Length > 0)
                    {
                        MessageHelper.Show(BexConstants.DataValidationTitle, validation.ToString(), MessageType.Stop);
                    }

                    var sb = new StringBuilder();
                    sb.AppendLine(ConvertToJson(packageModel.Map()));

                    packageModel.SegmentModels.ForEach(segmentModel =>
                    {
                        sb.AppendLine();
                        sb.AppendLine(ConvertToJson(segmentModel.Map()));

                        foreach (var hazardModel in segmentModel.HazardModels)
                        {
                            sb.AppendLine();
                            sb.AppendLine(ConvertToJson(hazardModel.Map()));
                        }

                        foreach (var policyModel in segmentModel.PolicyModels)
                        {
                            sb.AppendLine();
                            sb.AppendLine(ConvertToJson(policyModel.Map()));
                        }

                        foreach (var stateModel in segmentModel.StateModels)
                        {
                            sb.AppendLine();
                            sb.AppendLine(ConvertToJson(stateModel.Map()));
                        }

                        foreach (var exposureSetModel in segmentModel.ExposureSetModels)
                        {
                            sb.AppendLine();
                            sb.AppendLine(ConvertToJson(exposureSetModel.Map()));
                            foreach (var item in exposureSetModel.Items)
                            {
                                sb.AppendLine();
                                sb.AppendLine(ConvertToJson(item));
                            }
                        }

                        foreach (var aggregateLossSetModel in segmentModel.AggregateLossSetModels)
                        {
                            sb.AppendLine();
                            sb.AppendLine(ConvertToJson(aggregateLossSetModel.Map()));
                            foreach (var item in aggregateLossSetModel.Items)
                            {
                                sb.AppendLine();
                                sb.AppendLine(ConvertToJson(item));
                            }
                        }

                        foreach (var individualLossSetModel in segmentModel.IndividualLossSetModels)
                        {
                            sb.AppendLine();
                            sb.AppendLine(ConvertToJson(individualLossSetModel.Map()));
                            foreach (var item in individualLossSetModel.Items)
                            {
                                sb.AppendLine();
                                sb.AppendLine(ConvertToJson(item));
                            }
                        }

                    });

                    MessageHelper.Show("Create JSON", sb.ToString());
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("Build model failed", MessageType.Stop);
            }
        }

        public static void GetPackageFromBex(IWorkbookLogger logger)
        {
            try
            {
                var package = Globals.ThisWorkbook.ThisExcelWorkspace.Package;

                var packageId = package.SourceId;
                if (!packageId.HasValue)
                {
                    MessageHelper.Show("No package", MessageType.Stop);
                    return;
                }

                using (new CursorToWait())
                {
                    var sb = new StringBuilder();
                    using (new StatusBarUpdater("Getting package ..."))
                    {
                        var client = BexCollectorClientFactory.CreateBexCollectorClient(ConfigurationHelper.SecretWord,
                            ConfigurationHelper.UwpfTokenUrl, ConfigurationHelper.BexSubmissionsUrl, ConfigurationHelper.BexBaseUrl);
                        var task = client.SubmissionPackagesClient.GetAsync(packageId.Value);
                        task.Wait();

                        sb.AppendLine(ConvertToJson(task.Result));
                        sb.AppendLine();

                        var task2 = client.SubmissionSegmentsClient.GetAllAsync(packageId.Value);
                        task2.Wait();

                        var segments = task2.Result;
                        segments.ForEach(segment =>
                        {
                            sb.AppendLine(ConvertToJson(segment));
                            sb.AppendLine();

                            Debug.Assert(segment.Id.HasValue, $"{BexConstants.SegmentName} doesn't have an Id");
                            var task3 = client.HazardDistributionClient.GetAllAsync(packageId.Value, segment.Id.Value);
                            task3.Wait();
                            sb.AppendLine(ConvertToJson(task3.Result));
                            sb.AppendLine();

                            var task4 = client.PolicyDistributionClient.GetAllAsync(packageId.Value, segment.Id.Value);
                            task4.Wait();
                            sb.AppendLine(ConvertToJson(task4.Result));
                            sb.AppendLine();

                            var task5 = client.StateDistributionClient.GetAllAsync(packageId.Value, segment.Id.Value);
                            task5.Wait();
                            sb.AppendLine(ConvertToJson(task5.Result));
                            sb.AppendLine();

                            var task6 = client.ExposureLossSetsClient.GetAllAsync(packageId.Value, segment.Id.Value);
                            task6.Wait();
                            sb.AppendLine(ConvertToJson(task6.Result));
                            task6.Result.ForEach(exposureSet =>
                            {
                                Debug.Assert(exposureSet.Id.HasValue, $"{BexConstants.ExposureSetName} set doesn't have an Id");
                                var task7 = client.ExposureLossClient.GetAllAsync(packageId.Value, segment.Id.Value, exposureSet.Id.Value);
                                task7.Wait();
                                sb.AppendLine(ConvertToJson(task7.Result));
                            });
                            sb.AppendLine();

                            var task8 = client.AggregateLossSetsClient.GetAllAsync(packageId.Value, segment.Id.Value);
                            task8.Wait();
                            sb.AppendLine(ConvertToJson(task8.Result));
                            task8.Result.ForEach(aggregateLossSet =>
                            {
                                Debug.Assert(aggregateLossSet.Id.HasValue, $"{BexConstants.AggregateLossSetName} set doesn't have an Id");
                                var task9 = client.AggregateLossClient.GetAllAsync(packageId.Value, segment.Id.Value, aggregateLossSet.Id.Value);
                                task9.Wait();
                                sb.AppendLine(ConvertToJson(task9.Result));
                            });
                            sb.AppendLine();

                            var task10 = client.IndividualLossSetsClient.GetAllAsync(packageId.Value, segment.Id.Value);
                            task10.Wait();
                            sb.AppendLine(ConvertToJson(task10.Result));
                            task10.Result.ForEach(individualLossSet =>
                            {
                                Debug.Assert(individualLossSet.Id.HasValue, $"{BexConstants.IndividualLossSetName} Set doesn't have an Id");
                                var task11 = client.IndividualLossClient.GetAllAsync(packageId.Value, segment.Id.Value, individualLossSet.Id.Value);
                                task11.Wait();
                                sb.AppendLine(ConvertToJson(task11.Result));
                            });
                            sb.AppendLine();
                        });
                    }

                    
                    MessageHelper.Show("Create JSON", sb.ToString());
                }
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("Build model failed", MessageType.Stop);
            }
        }
    }
}
