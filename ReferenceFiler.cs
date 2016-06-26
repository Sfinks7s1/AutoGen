namespace Auto
{
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    using System;

    public class ReferenceFiler : DwgFiler
    {
        public ObjectIdCollection HardPointerIds;
        public ObjectIdCollection SoftPointerIds;
        public ObjectIdCollection HardOwnershipIds;
        public ObjectIdCollection SoftOwnershipIds;

        public ReferenceFiler()
        {
            HardPointerIds = new ObjectIdCollection();
            SoftPointerIds = new ObjectIdCollection();
            HardOwnershipIds = new ObjectIdCollection();
            SoftOwnershipIds = new ObjectIdCollection();
        }

        public override ErrorStatus FilerStatus
        {
            get { return ErrorStatus.OK; }
            set { }
        }

        public override FilerType FilerType => FilerType.IdFiler;

        public override long Position => 0;

        public override IntPtr ReadAddress()
        {
            return new IntPtr();
        }

        public override byte[] ReadBinaryChunk()
        {
            return null;
        }

        public override bool ReadBoolean()
        {
            return true;
        }

        public override byte ReadByte()
        {
            return new byte();
        }

        public override void ReadBytes(byte[] value)
        {
        }

        public override double ReadDouble()
        {
            return 0.0;
        }

        public override Handle ReadHandle()
        {
            return new Handle();
        }

        public override ObjectId ReadHardOwnershipId()
        {
            return ObjectId.Null;
        }

        public override ObjectId ReadHardPointerId()
        {
            return ObjectId.Null;
        }

        public override short ReadInt16()
        {
            return 0;
        }

        public override int ReadInt32()
        {
            return 0;
        }

        public override long ReadInt64()
        {
            return 0;
        }

        public override Point2d ReadPoint2d()
        {
            return new Point2d();
        }

        public override Point3d ReadPoint3d()
        {
            return new Point3d();
        }

        public override Scale3d ReadScale3d()
        {
            return new Scale3d();
        }

        public override ObjectId ReadSoftOwnershipId()
        {
            return ObjectId.Null;
        }

        public override ObjectId ReadSoftPointerId()
        {
            return ObjectId.Null;
        }

        public override string ReadString()
        {
            return null;
        }

        public override ushort ReadUInt16()
        {
            return 0;
        }

        public override uint ReadUInt32()
        {
            return 0;
        }

        public override ulong ReadUInt64()
        {
            return 0;
        }

        public override Vector2d ReadVector2d()
        {
            return new Vector2d();
        }

        public override Vector3d ReadVector3d()
        {
            return new Vector3d();
        }

        public override void ResetFilerStatus()
        {
        }

        public override void Seek(long offset, int method)
        {
        }

        public override void WriteAddress(IntPtr value)
        {
        }

        public override void WriteBinaryChunk(byte[] chunk)
        {
        }

        public override void WriteBoolean(bool value)
        {
        }

        public override void WriteByte(byte value)
        {
        }

        public override void WriteBytes(byte[] value)
        {
        }

        public override void WriteDouble(double value)
        {
        }

        public override void WriteHandle(Handle handle)
        {
        }

        public override void WriteInt16(short value)
        {
        }

        public override void WriteInt32(int value)
        {
        }

        public override void WriteInt64(long value)
        {
        }

        public override void WritePoint2d(Point2d value)
        {
        }

        public override void WritePoint3d(Point3d value)
        {
        }

        public override void WriteScale3d(Scale3d value)
        {
        }

        public override void WriteString(string value)
        {
        }

        public override void WriteUInt16(ushort value)
        {
        }

        public override void WriteUInt32(uint value)
        {
        }

        public override void WriteUInt64(ulong value)
        {
        }

        public override void WriteVector2d(Vector2d value)
        {
        }

        public override void WriteVector3d(Vector3d value)
        {
        }

        public override void WriteHardOwnershipId(ObjectId value)
        {
            HardOwnershipIds.Add(value);
        }

        public override void WriteHardPointerId(ObjectId value)
        {
            HardPointerIds.Add(value);
        }

        public override void WriteSoftOwnershipId(ObjectId value)
        {
            SoftOwnershipIds.Add(value);
        }

        public override void WriteSoftPointerId(ObjectId value)
        {
            SoftPointerIds.Add(value);
        }

        public void Reset()
        {
            HardPointerIds.Clear();

            SoftPointerIds.Clear();

            HardOwnershipIds.Clear();

            SoftOwnershipIds.Clear();
        }
    }
}