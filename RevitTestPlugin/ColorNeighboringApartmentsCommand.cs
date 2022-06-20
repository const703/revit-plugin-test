using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;

namespace RevitTestPlugin
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ColorNeighboringApartmentsCommand : IExternalCommand
    {
        private const string ZoneParameterName = "ROM_Зона";
        private const string BlockParameterName = "BS_Блок";
        private const string SubZoneParameterName = "ROM_Подзона";
        private const string SubZoneIdParameterName = "ROM_Расчетная_подзона_ID";
        private const string SubZoneIndexParameterName = "ROM_Подзона_Index";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                if (!ColorNeighboringApartments(commandData.Application.ActiveUIDocument.Document))
                    return Result.Failed;
            }
            catch
            {
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private static bool ColorNeighboringApartments(Document document)
        {
            var roomsCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Rooms);

            var firstRoom = roomsCollector.FirstElement();
            if (firstRoom == null)
                return true;
            
            var zoneParameter = firstRoom.LookupParameter(ZoneParameterName);
            var zoneParameterId = zoneParameter.Id;
            var zoneParameterGuid = zoneParameter.GUID;
            var blockParameterGuid = firstRoom.LookupParameter(BlockParameterName).GUID;
            var subZoneParameterGuid = firstRoom.LookupParameter(SubZoneParameterName).GUID;
            var subZoneIdParameterGuid = firstRoom.LookupParameter(SubZoneIdParameterName).GUID;
            var subZoneIndexParameterGuid = firstRoom.LookupParameter(SubZoneIndexParameterName).GUID;

            var roomsFilterRule = new FilterStringRule(new ParameterValueProvider(zoneParameterId), new FilterStringContains(), "Квартира", caseSensitive: true);

            var roomsFilter = new ElementParameterFilter(roomsFilterRule);

            roomsCollector.WherePasses(roomsFilter);

            var rooms = roomsCollector.ToElements();
            
            using (var transaction = new Transaction(document, "Color neighboring apartments"))
            {
                if (transaction.Start() != TransactionStatus.Started)
                    return false;

                foreach (var levelRooms in rooms.GroupBy(x => x.get_Parameter(BuiltInParameter.ROOM_LEVEL_ID).AsValueString()))
                {
                    foreach (var blockRooms in levelRooms.GroupBy(x => x.get_Parameter(blockParameterGuid).AsString()))
                    {
                        var roomsByApartment = blockRooms
                            .GroupBy(x => x.get_Parameter(zoneParameterGuid).AsString())
                            .OrderBy(x => x.Key);

                        string previousApartmentRoomsParameter = null;

                        foreach (var apartmentRooms in roomsByApartment)
                        {
                            var roomsNumberParameterValue = apartmentRooms.First().get_Parameter(subZoneParameterGuid).AsString();

                            if (!string.IsNullOrWhiteSpace(previousApartmentRoomsParameter) && previousApartmentRoomsParameter == roomsNumberParameterValue)
                            {
                                foreach (var room in apartmentRooms)
                                {
                                    var subZoneId = room.get_Parameter(subZoneIdParameterGuid).AsString();
                                    room.get_Parameter(subZoneIndexParameterGuid).Set(subZoneId + ".Полутон");
                                }

                                previousApartmentRoomsParameter = null;
                            }
                            else
                            {
                                previousApartmentRoomsParameter = roomsNumberParameterValue;
                            }
                        }
                    }
                }

                if (transaction.Commit() == TransactionStatus.Committed)
                    return true;

                transaction.RollBack();
            }

            return false;
        }
    }
}
