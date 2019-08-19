namespace mpFormats
{
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using ModPlusAPI;

    public class BlockJig : EntityJig
    {
        private const string LangItem = "mpFormats";
        Point3d _mCenterPt, _mActualPoint;

        public BlockJig(BlockReference br)
            : base(br)
        {
            _mCenterPt = br.Position;
        }

        public Vector3d ReplaceVector3D { get; private set; }

        protected override SamplerStatus Sampler(JigPrompts prompts)
        {
            var jigOpts = new JigPromptPointOptions
            {
                UserInputControls = (UserInputControls.Accept3dCoordinates
                                     | UserInputControls.NoZeroResponseAccepted
                                     | UserInputControls.AcceptOtherInputString
                                     | UserInputControls.NoNegativeResponseAccepted),
                Message = "\n" + Language.GetItem(LangItem, "msg11")
            };
            var dres = prompts.AcquirePoint(jigOpts);
            if (_mActualPoint == dres.Value)
                return SamplerStatus.NoChange;
            _mActualPoint = dres.Value;
            return SamplerStatus.OK;
        }

        protected override bool Update()
        {
            try
            {
                ((BlockReference)Entity).Position = _mActualPoint;
                ReplaceVector3D = _mCenterPt.GetVectorTo(_mActualPoint);

            }
            catch (System.Exception)
            {
                return false;
            }
            return true;
        }

        public Entity GetEntity()
        {
            return Entity;
        }

    }
}