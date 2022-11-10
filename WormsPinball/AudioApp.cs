namespace Windows
{
    using System;
    using System.Runtime.InteropServices;

    internal class AudioApp: IDisposable
    {
        private readonly ISimpleAudioVolume simpleVolume;
        private bool disposedValue;

        public string Name { get; private set; }
        public int ID { get; private set; }

        public bool Mute
        {
            get
            {
                if (simpleVolume != null)
                {
                    simpleVolume.GetMute(out bool mute);
                    return mute;
                }

                return false;
            }
            set
            {
                if (simpleVolume != null)
                {
                    var guid = Guid.Empty;
                    simpleVolume.SetMute(value, guid);
                }
            }
        }

        public float Volume
        {
            get
            {
                if (simpleVolume != null)
                {
                    simpleVolume.GetMasterVolume(out float level);
                    return level * 100;
                }

                return 100;
            }
            set
            {
                if (value >= 0 && value <= 100)
                {
                    if (simpleVolume != null)
                    {
                        var guid = Guid.Empty;
                        simpleVolume.SetMasterVolume(value / 100, guid);
                    }
                }
            }
        }

        public AudioApp(string name, int iD)
        {
            Name = name;
            ID = iD;
            if (ID > -1)
                simpleVolume = AudioManager.GetVolumeInterface((uint)ID);
        }

        ~AudioApp()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    //No managed objects to dispose
                }

                if (simpleVolume != null)
                    Marshal.ReleaseComObject(simpleVolume);

                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}
